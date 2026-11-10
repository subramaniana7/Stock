"""Independent daily NSE F&O-stock cash OHLC collection and weekly daily workbooks."""
import argparse, datetime as dt, hashlib, json, logging, logging.handlers, math
import pathlib, shutil, subprocess, sys, zipfile, xml.etree.ElementTree as ET
import update_weekly_ohlc as nse

BASE=pathlib.Path(__file__).resolve().parent
IST=dt.timezone(dt.timedelta(hours=5,minutes=30))
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def atomic_json(path,data):
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_suffix('.tmp')
    tmp.write_text(json.dumps(data,indent=2),encoding='utf-8'); tmp.replace(path)

def monday_of(day): return day-dt.timedelta(days=day.weekday())

def workbook_path(root,monday):
    return root/'outputs'/f'NSE_FO_Daily_OHLC_{monday}_to_{monday+dt.timedelta(days=4)}.xlsx'

def sheet_name(day): return day.strftime('%d-%b-%Y')

def due_dates(first,now,holidays,extra_sessions=(),extra_closures=()):
    result=[]; cursor=first
    while cursor<=now.date():
        session=(cursor.weekday()<5 and cursor not in holidays) or cursor in extra_sessions
        if session and cursor not in extra_closures and (cursor<now.date() or now.time().replace(tzinfo=None)>=dt.time(16,30)):
            result.append(cursor)
        cursor+=dt.timedelta(days=1)
    return result

def init_sources(root):
    # Only this Python process redirects the imported helpers to a private cache.
    # Existing weekly updater/configuration files and their caches are not written.
    nse.BASE=root/'daily_data'; nse.BASE.mkdir(exist_ok=True)
    for dirname in ['source_archives','calendars']:
        dest=nse.BASE/dirname; dest.mkdir(exist_ok=True)
        old=root/dirname
        if old.is_dir():
            for p in old.iterdir():
                if p.is_file() and p.suffix in {'.zip','.json'} and not (dest/p.name).exists():
                    shutil.copyfile(p,dest/p.name)

def load_calendar(root,year,now,offline=False):
    cache=root/'daily_data'/'calendars'/f'nse_trading_holidays_{year}.json'
    refreshed=False
    if cache.exists():
        data=json.loads(cache.read_text(encoding='utf-8'))
        refreshed=dt.datetime.fromisoformat(data['retrieved_at']).astimezone(IST).date()==now.date()
    return nse.trading_holidays(year,offline=offline or refreshed)

def collect_day(day,offline=False):
    cm,cm_source=nse.archive(day,'CM',offline)
    fo,fo_source=nse.archive(day,'FO',offline)
    symbols={r['TckrSymb'].strip() for r in fo if r['FinInstrmTp'] in {'STF','STO'}}
    if len(symbols)<100: raise ValueError(f'Unexpected F&O stock count: {len(symbols)}')
    cash={}
    for row in cm:
        if row['SctySrs'].strip()!='EQ' or row['SsnId'].strip()!='F1': continue
        symbol=row['TckrSymb'].strip()
        if symbol in cash: raise ValueError(f'Duplicate cash row: {symbol}')
        cash[symbol]=row
    prices=[]
    for symbol in sorted(symbols):
        if symbol not in cash: raise ValueError(f'Missing EQ/F1 cash price for {symbol}')
        row=cash[symbol]
        values=[float(row[k]) for k in ['OpnPric','HghPric','LwPric','ClsPric']]
        o,h,l,c=values
        if not all(math.isfinite(v) and v>0 for v in values) or not l<=min(o,c)<=max(o,c)<=h:
            raise ValueError(f'Invalid prices for {symbol}: {values}')
        prices.append([symbol,*values])
    return {'date':str(day),'rows':prices,'sources':[cm_source,fo_source],
            'retrieved_at':dt.datetime.now(IST).isoformat()}

def workbook_valid(path,days):
    if not path.exists(): return False
    try:
        with zipfile.ZipFile(path) as z:
            if z.testzip(): return False
            xml=ET.fromstring(z.read('xl/workbook.xml'))
            names={s.attrib['name'] for s in xml.findall('s:sheets/s:sheet',NS)}
        return names=={sheet_name(dt.date.fromisoformat(k)) for k in days}|{'Sources','CPR Guide'}
    except (OSError,zipfile.BadZipFile,KeyError,ET.ParseError): return False

def build_book(root,config,payload,render_dir=None):
    monday=dt.date.fromisoformat(payload['week_start'])
    output=workbook_path(root,monday); output.parent.mkdir(exist_ok=True)
    temporary=output.with_suffix('.pending.xlsx')
    if temporary.exists(): temporary.unlink()
    job=root/'daily_data'/'build_job.json'
    atomic_json(job,{'payload':payload,'output':str(temporary),'render_dir':str(render_dir) if render_dir else None})
    command=[config['node_executable'],str(config.get('builder_script',root/'daily_runtime'/'build_daily.mjs')),str(job)]
    try:
        result=subprocess.run(command,cwd=root,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=240,
                              creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        if not workbook_valid(temporary,payload['days']):
            raise RuntimeError(f'Daily Excel export failed ({result.returncode}): {result.stdout[-2000:]} {result.stderr[-2000:]}')
        # On Windows this fails safely if the workbook is open in Excel.
        temporary.replace(output)
    finally:
        if temporary.exists(): temporary.unlink()
    return output

def run_updates(root,config,now,logger,offline=False,render_dir=None,collector=None,builder=None):
    first=dt.date.fromisoformat(config['first_date'])
    holidays=set(); calendar_sources=[]
    for year in range(first.year,now.year+1):
        closed,info=load_calendar(root,year,now,offline)
        holidays.update(closed); calendar_sources.append(info)
    extra={dt.date.fromisoformat(x) for x in config.get('additional_sessions',[])}
    closures={dt.date.fromisoformat(x) for x in config.get('additional_closures',[])}
    due=due_dates(first,now,holidays,extra,closures)
    grouped={}
    for day in due: grouped.setdefault(monday_of(day),[]).append(day)
    collector=collector or collect_day; builder=builder or build_book
    failures=[]; added=[]
    for monday,days in sorted(grouped.items()):
        store=root/'daily_data'/'weeks'/f'{monday}.json'
        payload=json.loads(store.read_text(encoding='utf-8')) if store.exists() else {'week_start':str(monday),'days':{}}
        before=set(payload['days']); new=[]
        for day in days:
            if str(day) in payload['days']: continue
            try:
                payload['days'][str(day)]=collector(day,offline)
                new.append(str(day))
            except Exception as ex:
                failures.append(str(day)); logger.warning('PENDING %s: %s',day,ex)
        path=workbook_path(root,monday)
        if not payload['days']: continue
        if not new and workbook_valid(path,payload['days']): continue
        payload['calendar_sources']=calendar_sources
        try:
            path=builder(root,config,payload,render_dir)
            atomic_json(store,payload)  # Commit completion only after the Excel file is installed.
            added.extend(new)
            logger.info('SAVED %s; added dates: %s',path.name,', '.join(new) or 'restored workbook')
        except Exception as ex:
            failures.extend(new or [str(d) for d in days]); logger.exception('Workbook pending: %s',ex)
    return {'status':'pending' if failures else 'success','at':now.isoformat(),
            'added_dates':added,'pending_dates':sorted(set(failures)),
            'completed_days_due':len(due)-len(set(failures)),'first_date':str(first)}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--offline',action='store_true')
    parser.add_argument('--check-only',action='store_true')
    parser.add_argument('--render-dir',type=pathlib.Path)
    args=parser.parse_args()
    config=json.loads((BASE/'daily_config.json').read_text(encoding='utf-8'))
    now=dt.datetime.now(IST)
    logs=BASE/'logs'; logs.mkdir(exist_ok=True)
    logger=logging.getLogger('nse-daily'); logger.setLevel(logging.INFO)
    handler=logging.handlers.RotatingFileHandler(logs/'daily_update.log',maxBytes=2_000_000,backupCount=4,encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s')); logger.addHandler(handler)
    if args.check_only:
        print(json.dumps({'config':config,'now':now.isoformat()},indent=2)); return 0
    import msvcrt
    with (BASE/'daily_update.lock').open('a+b') as lock:
        lock.seek(0); lock.write(b'0'); lock.flush(); lock.seek(0)
        try: msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
        except OSError:
            logger.info('Another daily updater is active.'); return 0
        try:
            init_sources(BASE)
            status=run_updates(BASE,config,now,logger,args.offline,args.render_dir)
            atomic_json(BASE/'daily_last_run_status.json',status)
            if status['added_dates'] or status['pending_dates']: logger.info('%s',json.dumps(status))
            if sys.stdout: print(json.dumps(status,indent=2))
            return 1 if status['pending_dates'] else 0
        except Exception as ex:
            logger.exception('Daily update failed')
            status={'status':'failed','at':now.isoformat(),'error':str(ex)}
            atomic_json(BASE/'daily_last_run_status.json',status)
            if sys.stderr: print(str(ex),file=sys.stderr)
            return 1
        finally:
            lock.seek(0); msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1)

if __name__=='__main__': sys.exit(main())
