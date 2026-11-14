"""Monthly NSE F&O-stock cash OHLC, calendar-month catch-up and email notifications."""
import argparse,calendar,concurrent.futures,datetime as dt,hashlib,json,logging,logging.handlers,math
import pathlib,shutil,subprocess,sys,zipfile
import update_weekly_ohlc as nse
from email_notifications import queue_report,flush_notifications

BASE=pathlib.Path(__file__).resolve().parent
IST=dt.timezone(dt.timedelta(hours=5,minutes=30))
DEFAULT_CONFIG={'task_name':'NSE Monthly OHLC','first_month':'2026-10','additional_sessions':['2026-11-08'],'additional_closures':[]}

def month_start(value):
    return dt.datetime.strptime(value,'%Y-%m').date().replace(day=1)

def next_month(day):
    return dt.date(day.year+(day.month==12),1 if day.month==12 else day.month+1,1)

def previous_month(today):
    return (today.replace(day=1)-dt.timedelta(days=1)).replace(day=1)

def due_months(first_month,now):
    current=month_start(first_month);result=[]
    while dt.datetime.combine(next_month(current),dt.time(9),IST)<=now:
        result.append(current);current=next_month(current)
    return result

def atomic_json(path,data):
    path.parent.mkdir(parents=True,exist_ok=True);temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps(data,indent=2),encoding='utf-8');temp.replace(path)

def report_path(root,start):
    return root/'outputs'/f'NSE_FO_Monthly_OHLC_{start:%Y-%m}.xlsx'

def init_cache(root):
    nse.BASE=root/'monthly_data';nse.BASE.mkdir(exist_ok=True)
    for name in ['calendars','source_archives']:
        dest=nse.BASE/name;dest.mkdir(exist_ok=True)
        old=root/name
        if old.is_dir():
            for path in old.iterdir():
                if path.is_file() and path.suffix in {'.zip','.json'} and not (dest/path.name).exists():
                    shutil.copyfile(path,dest/path.name)

def session_dates(start,holidays,extras=(),closures=()):
    end=next_month(start);extra=set(extras);closed=set(closures)
    return [start+dt.timedelta(days=i) for i in range((end-start).days)
        if ((start+dt.timedelta(days=i)).weekday()<5 and start+dt.timedelta(days=i) not in holidays
            or start+dt.timedelta(days=i) in extra) and start+dt.timedelta(days=i) not in closed]

def aggregate(start,days,cash,universes,sources,calendar_info):
    symbols=set().union(*universes.values());daily=[];rows=[];calculated=[]
    # Merge symbol changes only when the official cash files share the same ISIN.
    parent={s:s for s in symbols};by_isin={}
    def find(s):
        while parent[s]!=s:parent[s]=parent[parent[s]];s=parent[s]
        return s
    for symbol in sorted(symbols):
        for day in days:
            isin=cash[day].get(symbol,{}).get('ISIN')
            if not isin:continue
            if isin in by_isin:parent[find(symbol)]=find(by_isin[isin])
            else:by_isin[isin]=symbol
    groups={}
    for symbol in symbols:groups.setdefault(find(symbol),set()).add(symbol)
    records=[];mapping=[]
    for aliases in groups.values():
        canonical=max(aliases,key=lambda s:(max((d for d in days if s in cash[d]),default=dt.date.min),s))
        isins={cash[d][s].get('ISIN') for d in days for s in aliases if s in cash[d]}-{None,''}
        records.append((canonical,aliases,isins))
        if len(aliases)>1:mapping.append(['Symbol mapping',', '.join(sorted(aliases))+' -> '+canonical,', '.join(sorted(isins)),'Matched using official cash-file ISIN'])
    for symbol,aliases,isins in sorted(records):
        first=len(daily)+2;prices=[]
        for day in days:
            matches=[(s,r) for s,r in cash[day].items() if s in aliases or r.get('ISIN') in isins]
            if len(matches)!=1:raise ValueError(f'Missing or ambiguous monthly cash observation: {symbol} on {day}; no partial report exported')
            source_symbol,r=matches[0];values=[float(r[k]) for k in ['OpnPric','HghPric','LwPric','ClsPric']]
            o,h,l,c=values
            if not all(math.isfinite(v) and v>0 for v in values) or not l<=min(o,c)<=max(o,c)<=h:
                raise ValueError(f'Invalid cash OHLC: {symbol} on {day}')
            daily.append([symbol,str(day),*values,'Yes' if aliases.intersection(universes[day]) else 'No',source_symbol]);prices.append(values)
        last=len(daily)+1
        rows.append([symbol,f"='Daily Prices'!C{first}",f"=MAX('Daily Prices'!D{first}:D{last})",f"=MIN('Daily Prices'!E{first}:E{last})",f"='Daily Prices'!F{last}"])
        calculated.append([symbol,prices[0][0],max(r[1] for r in prices),min(r[2] for r in prices),prices[-1][3]])
    end=next_month(start)-dt.timedelta(days=1)
    notes=[['Report month',f'{start:%B %Y} ({start} to {end})'],['Currency','INR per share; unadjusted NSE cash EQ/F1 prices'],
        ['Universe',f'{len(records)} unique stocks from {len(symbols)} underlying symbols; ISIN-confirmed renames consolidated; index derivatives excluded'],
        ['Sessions',', '.join(str(d) for d in days)],['Method','First session open; maximum daily high; minimum daily low; final session close'],
        ['Price coverage','Every included stock requires a cash observation on every selected monthly session'],
        ['F&O eligibility','Union of that month\'s derivatives files; daily eligibility flags retained'],
        ['Corporate actions','Prices are not adjusted for splits, dividends or other corporate actions'],
        ['Special sessions','Weekdays excluding NSE closures, plus explicitly confirmed additional sessions in monthly_config.json'],
        ['Holiday source',calendar_info['source']],['Fetched at',dt.datetime.now(IST).isoformat()],[],
        ['Date','Segment','Official source URL','SHA-256'],*sources,*mapping]
    return {'month':start.strftime('%Y-%m'),'start':str(start),'end':str(end),'sessions':[str(d) for d in days],
            'count':len(records),'rows':rows,'calculated':calculated,'daily':daily,'notes':notes,'symbol_mappings':mapping}

def prepare_month(start,config,logger,offline=False):
    holidays,info=nse.trading_holidays(start.year,offline)
    days=session_dates(start,holidays,[dt.date.fromisoformat(x) for x in config.get('additional_sessions',[])],
        [dt.date.fromisoformat(x) for x in config.get('additional_closures',[])])
    if not days:raise ValueError('No trading sessions selected for this month')
    cash={};universes={};sources=[]
    def fetch(day):
        cm,cm_source=nse.archive(day,'CM',offline);fo,fo_source=nse.archive(day,'FO',offline)
        symbols={r['TckrSymb'].strip() for r in fo if r['FinInstrmTp'] in {'STF','STO'}}
        if len(symbols)<100:raise ValueError(f'Unexpected F&O universe on {day}: {len(symbols)} stocks')
        observations={}
        for row in cm:
            if row['SctySrs'].strip()!='EQ' or row['SsnId'].strip()!='F1':continue
            symbol=row['TckrSymb'].strip()
            if symbol in observations:raise ValueError(f'Duplicate cash key: {symbol} on {day}')
            observations[symbol]={k:row[k] for k in ['OpnPric','HghPric','LwPric','ClsPric','ISIN']}
        logger.info('Retrieved and validated sources for %s',day)
        if sys.stdout:print(f'Retrieved {day}: cash + F&O files',flush=True)
        return day,observations,symbols,[cm_source,fo_source]
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
        downloaded=list(pool.map(fetch,days))
    for day,observations,symbols,day_sources in downloaded:
        universes[day]=symbols;cash[day]=observations
        sources.extend(day_sources);logger.info('Validated CM/FO sources: %s (%s F&O stocks)',day,len(symbols))
        if sys.stdout:print(f'Validated {day}: cash + F&O files, {len(symbols)} stocks',flush=True)
    return aggregate(start,days,cash,universes,sources,info)

def export_month(root,payload,render_dir=None):
    config=json.loads((root/'daily_config.json').read_text(encoding='utf-8'))
    output=report_path(root,month_start(payload['month']));output.parent.mkdir(exist_ok=True)
    temp=output.with_suffix('.pending.xlsx');job=root/'monthly_data'/'build_job.json'
    if temp.exists():temp.unlink()
    atomic_json(job,{'payload':payload,'output':str(temp),'render_dir':str(render_dir) if render_dir else None})
    try:
        result=subprocess.run([config['node_executable'],str(root/'daily_runtime/build_monthly.mjs'),str(job)],
            cwd=root,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=300,
            creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        if not temp.exists():raise RuntimeError(f'Monthly workbook export failed: {result.stdout[-1500:]} {result.stderr[-1500:]}')
        with zipfile.ZipFile(temp) as z:
            if z.testzip() or 'xl/workbook.xml' not in z.namelist():raise ValueError('Invalid monthly workbook')
        temp.replace(output)
    finally:
        if temp.exists():temp.unlink()
    return output

def run_months(root,months,config,now,logger,offline=False,fetch_only=False,render_dir=None):
    state_path=root/'monthly_state.json';state=json.loads(state_path.read_text()) if state_path.exists() else {'completed':{}}
    failed=[];saved=[]
    for start in months:
        key=start.strftime('%Y-%m');output=report_path(root,start);receipt=state['completed'].get(key)
        if not fetch_only and receipt and output.exists() and hashlib.sha256(output.read_bytes()).hexdigest()==receipt['sha256']:continue
        try:
            payload=prepare_month(start,config,logger,offline)
            atomic_json(root/'monthly_data'/'months'/f'{key}.json',payload)
            if fetch_only:continue
            output=export_month(root,payload,render_dir)
            state['completed'][key]={'file':output.name,'completed_at':now.isoformat(),'sha256':hashlib.sha256(output.read_bytes()).hexdigest()}
            atomic_json(state_path,state)
            queue_report(root,'monthly:'+key,output,key,'created',logger)
            saved.append(key);logger.info('SAVED monthly report: %s',output)
        except Exception as ex:
            failed.append(key);logger.exception('PENDING monthly report %s: %s',key,ex)
            if sys.stderr:print(f'Month {key} pending: {ex}',file=sys.stderr,flush=True)
    notification=flush_notifications(root,logger) if not fetch_only else {'sent':0,'pending':0}
    return {'status':'pending' if failed else 'success','at':now.isoformat(),'saved_months':saved,'pending_months':failed,'email_notifications':notification}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--month',help='Completed calendar month YYYY-MM; default all due months')
    parser.add_argument('--offline',action='store_true');parser.add_argument('--fetch-only',action='store_true')
    parser.add_argument('--render-dir',type=pathlib.Path)
    args=parser.parse_args();now=dt.datetime.now(IST)
    path=BASE/'monthly_config.json';config=json.loads(path.read_text()) if path.exists() else DEFAULT_CONFIG
    months=[month_start(args.month)] if args.month else due_months(config['first_month'],now)
    if any(next_month(d)>now.date().replace(day=1) for d in months):parser.error('Only completed calendar months are supported')
    logs=BASE/'logs';logs.mkdir(exist_ok=True);logger=logging.getLogger('nse-monthly');logger.setLevel(logging.INFO)
    handler=logging.handlers.RotatingFileHandler(logs/'monthly_update.log',maxBytes=2_000_000,backupCount=4,encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'));logger.addHandler(handler)
    import msvcrt
    with (BASE/'monthly_update.lock').open('a+b') as lock:
        if lock.seek(0,2)==0:lock.write(b'0');lock.flush()
        lock.seek(0)
        try:msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
        except OSError:return 0
        try:
            init_cache(BASE);logger.info('Monthly runner started: %s',now.isoformat())
            status=run_months(BASE,months,config,now,logger,args.offline,args.fetch_only,args.render_dir)
            atomic_json(BASE/'monthly_last_run_status.json',status)
            if sys.stdout:print(json.dumps(status,indent=2))
            return 1 if status['pending_months'] else 0
        finally:lock.seek(0);msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1)

if __name__=='__main__':sys.exit(main())
