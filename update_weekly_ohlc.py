"""NSE cash weekly OHLC for historical F&O stock universe. Python 3.10+, no pip packages."""
import argparse, csv, datetime as dt, hashlib, io, json, math, pathlib, sys, time
import urllib.request, zipfile
from xml.sax.saxutils import escape
from email_notifications import queue_report,flush_notifications

BASE = pathlib.Path(__file__).resolve().parent
HEADERS = ['Stock Symbol','Weekly Open','Weekly High','Weekly Low','Weekly Close']

def previous_week(today):
    # On Saturday/Sunday the current Monday-Friday trading week has finished.
    monday = today - dt.timedelta(days=today.weekday()+(0 if today.weekday()>=5 else 7))
    return monday, monday+dt.timedelta(days=4)

def trading_holidays(year, offline=False):
    """Use only a validated NSE CM/FO calendar; never infer closures from 404s."""
    url='https://www.nseindia.com/api/holiday-master?type=trading'
    cache=BASE/'calendars'/f'nse_trading_holidays_{year}.json'
    failure=None
    if not offline:
        try:
            req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0','Referer':'https://www.nseindia.com/'})
            raw=json.load(urllib.request.urlopen(req,timeout=30))
            parsed={}
            for segment in ['CM','FO']:
                entries=raw.get(segment)
                if not isinstance(entries,list) or not entries: raise ValueError(f'Missing {segment} holiday list')
                dates={dt.datetime.strptime(r['tradingDate'],'%d-%b-%Y').date() for r in entries}
                parsed[segment]=sorted(str(d) for d in dates if d.year==year)
                if len(parsed[segment])<5: raise ValueError(f'NSE response does not cover {year}')
            data={'year':year,'source':url,'retrieved_at':dt.datetime.now(dt.timezone.utc).isoformat(),'holidays':parsed}
            cache.parent.mkdir(exist_ok=True)
            temp=cache.with_suffix('.tmp'); temp.write_text(json.dumps(data,indent=2),encoding='utf-8'); temp.replace(cache)
        except Exception as ex:
            failure=ex
    if not cache.exists(): raise RuntimeError(f'Official NSE holiday calendar unavailable for {year}: {failure or "no cached calendar"}')
    data=json.loads(cache.read_text(encoding='utf-8'))
    if data.get('year')!=year or data.get('source')!=url: raise ValueError('Invalid cached holiday calendar')
    cm={dt.date.fromisoformat(s) for s in data['holidays']['CM']}
    fo={dt.date.fromisoformat(s) for s in data['holidays']['FO']}
    if len(cm)<5 or len(fo)<5 or any(d.year!=year for d in cm|fo): raise ValueError('Invalid cached holiday dates')
    if cm!=fo: raise ValueError('CM and FO holiday calendars differ; manual review required')
    if failure: print(f'WARNING: Using cached {year} NSE holiday calendar ({data["retrieved_at"]}): {failure}',flush=True)
    return cm,data

def archive(day, segment, offline=False):
    name=f'BhavCopy_NSE_{segment}_0_0_0_{day:%Y%m%d}_F_0000.csv.zip'
    url=f'https://nsearchives.nseindia.com/content/{segment.lower()}/{name}'
    p=BASE/'source_archives'/name
    if p.exists():
        data=p.read_bytes()
    elif offline:
        raise RuntimeError(f'Missing cached archive: {name}')
    else:
        error=None
        for attempt in range(3):
            try:
                req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0','Referer':'https://www.nseindia.com/'})
                data=urllib.request.urlopen(req,timeout=45).read()
                break
            except Exception as ex:
                error=ex
                if attempt<2: time.sleep(attempt+1)
        else:
            raise RuntimeError(f'Cannot download {day} {segment}: {error}. No partial workbook was written. If this is a confirmed exchange holiday, rerun with --holiday {day}.')
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names=[n for n in z.namelist() if n.lower().endswith('.csv')]
            if len(names)!=1: raise ValueError('Expected one CSV')
            rows=list(csv.DictReader(io.StringIO(z.read(names[0]).decode('utf-8-sig'))))
        required={'TradDt','TckrSymb','FinInstrmTp','OpnPric','HghPric','LwPric','ClsPric','SctySrs','SsnId'}
        if not rows or not required.issubset(rows[0]): raise ValueError('Missing required columns')
        if any(r['TradDt'][:10]!=day.isoformat() or r['Sgmt']!=segment for r in rows):
            raise ValueError('Archive date or segment does not match request')
    except Exception as ex:
        raise RuntimeError(f'Invalid NSE archive {name}: {ex}') from ex
    p.parent.mkdir(exist_ok=True)
    if not p.exists(): p.write_bytes(data)
    return rows,[day.isoformat(),segment,url,hashlib.sha256(data).hexdigest()]

def prepare(start,end,holidays=(),include=(),offline=False):
    days=[start+dt.timedelta(days=i) for i in range(7) if (start+dt.timedelta(days=i)).weekday()<5]
    days=sorted((set(days)-set(holidays))|set(include))
    if not days or any(d<start or d>start+dt.timedelta(days=6) for d in days): raise ValueError('Invalid trading-day selection')
    universe=set(); cash={}; sources=[]; daily_universe={}
    for day in days:
        cm,cm_src=archive(day,'CM',offline)
        fo,fo_src=archive(day,'FO',offline)
        symbols={r['TckrSymb'].strip() for r in fo if r['FinInstrmTp'] in {'STF','STO'}}
        if len(symbols)<100: raise ValueError(f'Unexpectedly small F&O universe on {day}: {len(symbols)}')
        universe.update(symbols); daily_universe[day]=symbols; sources.extend([cm_src,fo_src])
        cash[day]={}
        for r in cm:
            if r['SctySrs'].strip()!='EQ' or r['SsnId'].strip()!='F1': continue
            symbol=r['TckrSymb'].strip()
            if symbol in cash[day]: raise ValueError(f'Duplicate EQ/F1 cash record: {day} {symbol}')
            cash[day][symbol]=r
    daily=[]; weekly=[]; calculated=[]
    for symbol in sorted(universe):
        first=len(daily)+2
        prices=[]
        for day in days:
            if symbol not in cash[day]: raise ValueError(f'Missing cash EQ/F1 row: {symbol} on {day}; no partial workbook written')
            r=cash[day][symbol]
            values=[float(r[k]) for k in ['OpnPric','HghPric','LwPric','ClsPric']]
            o,h,l,c=values
            if not all(math.isfinite(v) and v>0 for v in values) or not l<=min(o,c)<=max(o,c)<=h:
                raise ValueError(f'Invalid cash prices: {symbol} on {day}: {values}')
            prices.append(values)
            daily.append([symbol,day.isoformat(),*values,'Yes' if symbol in daily_universe[day] else 'No'])
        last=len(daily)+1
        vals=[prices[0][0],max(x[1] for x in prices),min(x[2] for x in prices),prices[-1][3]]
        weekly.append([symbol,f"='Daily Prices'!C{first}",f"=MAX('Daily Prices'!D{first}:D{last})",f"=MIN('Daily Prices'!E{first}:E{last})",f"='Daily Prices'!F{last}"])
        calculated.append([symbol,*vals])
    notes=[['NSE weekly cash-market OHLC',f'{start} to {end}'],['Currency','INR per share; unadjusted NSE cash prices'],['Universe',f'{len(universe)} unique stock underlyings appearing in STF/STO contracts during selected sessions; excludes index contracts'],['Sessions',', '.join(str(x) for x in days)],['Method','First session open; maximum daily high; minimum daily low; final session close'],['Historical eligibility','Union of daily F&O bhavcopies; eligibility flag is retained for each cash observation'],['Price coverage','All selected cash sessions required for every stock; missing or invalid data stops export'],['Corporate actions','No split, dividend or other corporate-action adjustment is applied'],['Special sessions','Default Monday-Friday. Use --include-date for confirmed weekend sessions and --holiday for confirmed holidays.'],['Source files','Official NSE UDiFF final bhavcopy ZIPs; SHA-256 below'],['Fetched at',dt.datetime.now(dt.timezone(dt.timedelta(hours=5,minutes=30))).isoformat()],[],['Date','Segment','Source URL','SHA-256'],*sources]
    return {'start':str(start),'end':str(end),'weekly':weekly,'calculated':calculated,'daily':daily,'notes':notes,'count':len(universe)}

def column(i):
    s=''
    while i: i,n=divmod(i-1,26); s=chr(65+n)+s
    return s

def write_xlsx(path,payload):
    """Export with the installed workbook runtime, including live CPR formulas."""
    import subprocess
    path=pathlib.Path(path)
    config=json.loads((BASE/'daily_config.json').read_text(encoding='utf-8'))
    job=path.with_suffix('.job.json')
    job.write_text(json.dumps({'payload':payload,'output':str(path)}),encoding='utf-8')
    try:
        if path.exists(): path.unlink()
        result=subprocess.run([config['node_executable'],str(BASE/'daily_runtime/build_weekly.mjs'),str(job)],
            cwd=BASE,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=240,
            creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        if not path.exists(): raise RuntimeError(f'Weekly Excel export failed: {result.stdout[-1500:]} {result.stderr[-1500:]}')
        with zipfile.ZipFile(path) as z:
            if z.testzip() or 'xl/workbook.xml' not in z.namelist(): raise RuntimeError('Invalid weekly Excel export')
    finally:
        if job.exists(): job.unlink()


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--as-of',type=dt.date.fromisoformat,help='Reference date YYYY-MM-DD; default today in India')
    parser.add_argument('--week-start',type=dt.date.fromisoformat,help='Explicit Monday YYYY-MM-DD')
    parser.add_argument('--holiday',action='append',default=[],type=dt.date.fromisoformat,help='Confirmed exchange closure; repeat as needed')
    parser.add_argument('--include-date',action='append',default=[],type=dt.date.fromisoformat,help='Confirmed special session inside Monday-Sunday week')
    parser.add_argument('--offline',action='store_true',help='Require previously cached official archives')
    parser.add_argument('--auto-holidays',action='store_true',help='Exclude holidays from official NSE CM/FO calendars')
    parser.add_argument('--output-dir',type=pathlib.Path,help='Alternate output folder (default: outputs beside this script)')
    parser.add_argument('--json-only',action='store_true',help=argparse.SUPPRESS)
    args=parser.parse_args()
    today=args.as_of or dt.datetime.now(dt.timezone(dt.timedelta(hours=5,minutes=30))).date()
    start,end=previous_week(today)
    if args.week_start:
        start=args.week_start; end=start+dt.timedelta(days=4)
        if start.weekday()!=0: parser.error('--week-start must be a Monday')
        if start+dt.timedelta(days=5)>today: parser.error('Requested Monday-Friday trading week is not yet completed')
    if any(d<start or d>start+dt.timedelta(days=6) for d in args.holiday): parser.error('--holiday must lie in requested week')
    if any(d>=today for d in args.include_date): parser.error('--include-date must be a completed past session')
    calendar_sources=[]
    if args.auto_holidays:
        for year in sorted({start.year,(start+dt.timedelta(days=6)).year}):
            closed,info=trading_holidays(year,args.offline)
            args.holiday.extend(d for d in closed if start<=d<=start+dt.timedelta(days=6))
            calendar_sources.append(info)
    payload=prepare(start,end,args.holiday,args.include_date,args.offline)
    for info in calendar_sources:
        payload['notes'].append(['Holiday calendar',str(info['year']),info['source'],'Retrieved '+info['retrieved_at']])
    if args.holiday: payload['notes'].append(['Excluded closures',', '.join(str(d) for d in sorted(set(args.holiday)))])
    out=args.output_dir or BASE/'outputs'; out.mkdir(parents=True,exist_ok=True)
    stem=f'NSE_FO_Weekly_OHLC_{start}_{end}'
    (out/(stem+'.json')).write_text(json.dumps(payload,indent=2),encoding='utf-8')
    if not args.json_only:
        temp=out/(stem+'.tmp.xlsx'); final=out/(stem+'.xlsx')
        try: write_xlsx(temp,payload); temp.replace(final)
        finally:
            if temp.exists(): temp.unlink()
        queue_report(BASE,'weekly:'+str(start),final,f'{start} to {end}','created')
        flush_notifications(BASE)
        print(f'Saved {final} ({payload["count"]} stocks; {len(payload["daily"])} daily cash observations).')
    else: print(f'Validated {payload["count"]} stocks; {len(payload["daily"])} daily observations.')

if __name__=='__main__':
    try: main()
    except Exception as ex:
        print(f'ERROR: {ex}',file=sys.stderr); sys.exit(1)
