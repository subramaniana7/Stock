"""Current official Nifty 500 constituent OHLC reports, independent weekly/monthly jobs."""
import argparse,concurrent.futures,csv,datetime as dt,hashlib,io,json,logging,logging.handlers
import math,pathlib,re,shutil,subprocess,sys,urllib.request,uuid,zipfile
import update_weekly_ohlc as nse
from monthly_update import month_start,next_month,due_months,session_dates
from report_paths import month_folder,weekly_filename,MONTHS
from email_notifications import queue_report,flush_notifications

BASE=pathlib.Path(__file__).resolve().parent
IST=dt.timezone(dt.timedelta(hours=5,minutes=30))
URLS=['https://www.niftyindices.com/IndexConstituent/ind_nifty500list.csv','https://nsearchives.nseindia.com/content/indices/ind_nifty500list.csv']
DEFAULT_CONFIG={'first_week':'2026-11-09','first_month':'2026-10','additional_sessions':['2026-11-08'],'additional_closures':[]}

def atomic_json(path,value):
    path.parent.mkdir(parents=True,exist_ok=True);temp=path.with_name(path.name+'.'+uuid.uuid4().hex+'.tmp')
    temp.write_text(json.dumps(value,indent=2),encoding='utf-8');temp.replace(path)

def atomic_bytes(path,data):
    temp=path.with_name(path.name+'.'+uuid.uuid4().hex+'.tmp')
    temp.write_bytes(data);temp.replace(path)

def parse_constituents(data):
    rows=list(csv.DictReader(io.StringIO(data.decode('utf-8-sig'))))
    required={'Company Name','Industry','Symbol','Series','ISIN Code'}
    if not rows or not required.issubset(rows[0]):raise ValueError('Official Nifty 500 CSV schema is invalid')
    result=[];symbols=set()
    for row in rows:
        row={k:v.strip() for k,v in row.items() if k in required}
        symbol=row['Symbol'];isin=row['ISIN Code']
        if not re.fullmatch(r'[A-Z0-9&_.-]+',symbol) or not re.fullmatch(r'[A-Z0-9]{12}',isin):raise ValueError('Invalid constituent identifier')
        if symbol in symbols:raise ValueError('Duplicate constituent symbol: '+symbol)
        symbols.add(symbol);result.append(row)
    if not 490<=len(result)<=550:raise ValueError(f'Unexpected Nifty 500 entry count: {len(result)}')
    return sorted(result,key=lambda r:r['Symbol'])

def current_constituents(root,now,offline=False):
    folder=root/'nifty500_data';folder.mkdir(exist_ok=True)
    info_file=folder/'current_constituents.json'
    if not info_file.exists() and (folder/'initial_constituents.json').exists():
        info=json.loads((folder/'initial_constituents.json').read_text())
        data=(folder/'initial_constituents.csv').read_bytes();parse_constituents(data)
        if hashlib.sha256(data).hexdigest()!=info['sha256']:raise ValueError('Initial list checksum mismatch')
        filename='constituents_'+info['downloaded_at'][:10]+'_'+info['sha256'][:12]+'.csv'
        atomic_bytes(folder/filename,data);info['snapshot_file']=filename;atomic_json(info_file,info)
    info=json.loads(info_file.read_text()) if info_file.exists() else None
    fresh=info and dt.datetime.fromisoformat(info['downloaded_at']).astimezone(IST).date()==now.date()
    if not offline and not fresh:
        error=None
        for url in URLS:
            try:
                request=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0','Referer':'https://www.niftyindices.com/'})
                data=urllib.request.urlopen(request,timeout=30).read();rows=parse_constituents(data)
                sha=hashlib.sha256(data).hexdigest();filename=f'constituents_{now:%Y-%m-%d}_{sha[:12]}.csv'
                atomic_bytes(folder/filename,data)
                info={'source':url,'downloaded_at':now.isoformat(),'sha256':sha,'count':len(rows),'snapshot_file':filename}
                atomic_json(info_file,info);break
            except Exception as ex:error=ex
        else:raise RuntimeError(f'Fresh official constituent list unavailable: {error}')
    if not info:raise RuntimeError('No validated official constituent snapshot is cached')
    data=(folder/info['snapshot_file']).read_bytes()
    if hashlib.sha256(data).hexdigest()!=info['sha256']:raise ValueError('Constituent snapshot checksum mismatch')
    return parse_constituents(data),info

def init_cache(root):
    nse.BASE=root/'nifty500_data';nse.BASE.mkdir(exist_ok=True)
    for name in ['source_archives','calendars']:
        dest=nse.BASE/name;dest.mkdir(exist_ok=True)
        for parent in [root,root/'monthly_data',root/'daily_data']:
            source=parent/name
            if not source.is_dir() or source==dest:continue
            for item in source.iterdir():
                allowed=item.suffix=='.json' if name=='calendars' else '_CM_' in item.name and item.suffix=='.zip'
                if item.is_file() and allowed and not (dest/item.name).exists():shutil.copyfile(item,dest/item.name)

def report_path(root,kind,start):
    name='NSE500 '+weekly_filename(start) if kind=='weekly' else f'NSE500 Monthly {MONTHS[start.month-1]} {start.year}.xlsx'
    return month_folder(root/'outputs',start)/name

def pending_periods(kind,config,now):
    if kind=='monthly':return due_months(config['first_month'],now)
    day=dt.date.fromisoformat(config['first_week'])
    if day.weekday()!=0:raise ValueError('first_week must be a Monday')
    result=[]
    while dt.datetime.combine(day+dt.timedelta(days=5),dt.time(9),IST)<=now:
        result.append(day);day+=dt.timedelta(days=7)
    return result

def select_cash(member,rows):
    allowed={'EQ':0,'BE':1,'BZ':2,'RR':3}
    by_isin=[r for r in rows if r['SsnId']=='F1' and r['SctySrs'] in allowed and r['ISIN']==member['ISIN Code']]
    matches=by_isin or [r for r in rows if r['SsnId']=='F1' and r['SctySrs'] in allowed and r['TckrSymb']==member['Symbol']]
    if not matches:return None
    rank=min(allowed[r['SctySrs']] for r in matches);matches=[r for r in matches if allowed[r['SctySrs']]==rank]
    if len(matches)!=1:raise ValueError('Ambiguous cash observation for '+member['Symbol'])
    row=matches[0];values=[float(row[k]) for k in ['OpnPric','HghPric','LwPric','ClsPric']]
    o,h,l,c=values
    if not all(math.isfinite(v) and v>0 for v in values) or not l<=min(o,c)<=max(o,c)<=h:
        raise ValueError('Invalid cash OHLC for '+member['Symbol'])
    return [*values,row['TckrSymb'],row['SctySrs']]

def aggregate(kind,start,days,members,info,cash,sources,calendar_info):
    daily=[];rows=[];calculated=[];coverage=[]
    for member in members:
        symbol=member['Symbol'];first=len(daily)+2;observations=[]
        for day in days:
            observation=select_cash(member,cash[day])
            if observation is None:continue
            observations.append((day,observation));daily.append([symbol,str(day),*observation])
        last=len(daily)+1;complete=len(observations)==len(days)
        status='Complete' if complete else ('Unavailable: no cash prices in period' if not observations else 'Unavailable: incomplete period coverage')
        coverage.append([symbol,status,len(observations),len(days),str(observations[0][0]) if observations else '',str(observations[-1][0]) if observations else '',member['ISIN Code']])
        if complete:
            prices=[r for _,r in observations]
            rows.append([symbol,f"='Daily Prices'!C{first}",f"=MAX('Daily Prices'!D{first}:D{last})",f"=MIN('Daily Prices'!E{first}:E{last})",f"='Daily Prices'!F{last}"])
            calculated.append([symbol,prices[0][0],max(r[1] for r in prices),min(r[2] for r in prices),prices[-1][3]])
        else:rows.append([symbol,'n.a.','n.a.','n.a.','n.a.']);calculated.append([symbol,None,None,None,None])
    end=start+dt.timedelta(days=4) if kind=='weekly' else next_month(start)-dt.timedelta(days=1)
    notes=[['Report',f'NSE500 {kind}: {start} to {end}'],['Prices','Unadjusted NSE cash-market OHLC, INR per share'],
        ['Universe',f'CURRENT official Nifty 500 list: {len(members)} entries; membership is not historical'],
        ['Constituents fetched','IST '+info['downloaded_at']],['Constituent source',info['source']],['Constituent SHA-256',info['sha256']],
        ['Sessions',', '.join(str(d) for d in days)],['Cash matching','ISIN preferred; symbol fallback; EQ, BE, BZ or RR regular-series records; original symbols/series retained'],
        ['Unavailable data','All index entries retained. Missing or incomplete-period prices are n.a.; see Coverage. No prices are invented.'],
        ['Holiday source',calendar_info['source']],['Method','First session open; maximum high; minimum low; final session close'],[],
        ['Date','Segment','Official cash-source URL','SHA-256'],*sources]
    return {'kind':kind,'start':str(start),'end':str(end),'count':len(members),'rows':rows,'calculated':calculated,'daily':daily,
            'coverage':coverage,'members':members,'constituents':info,'notes':notes,'sessions':[str(d) for d in days]}

def prepare(root,kind,start,config,now,logger,offline=False):
    members,info=current_constituents(root,now,offline);holidays,calendar_info=nse.trading_holidays(start.year,offline)
    if kind=='weekly' and (start+dt.timedelta(days=4)).year!=start.year:
        next_holidays,next_info=nse.trading_holidays((start+dt.timedelta(days=4)).year,offline)
        holidays=set(holidays)|set(next_holidays)
        calendar_info={'source':calendar_info['source']+'; '+next_info['source']}
    if kind=='monthly':days=session_dates(start,holidays,[dt.date.fromisoformat(d) for d in config.get('additional_sessions',[])],[dt.date.fromisoformat(d) for d in config.get('additional_closures',[])])
    else:days=[start+dt.timedelta(days=i) for i in range(5) if start+dt.timedelta(days=i) not in holidays and str(start+dt.timedelta(days=i)) not in config.get('additional_closures',[])]
    if not days:raise ValueError('No sessions selected')
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:data=list(pool.map(lambda day:(day,*nse.archive(day,'CM',offline)),days))
    cash={day:observations for day,observations,_ in data};sources=[source for _,_,source in data]
    payload=aggregate(kind,start,days,members,info,cash,sources,calendar_info)
    logger.info('Validated %s: %s members, %s sessions, %s unavailable',kind,len(members),len(days),sum(r[1]!='Complete' for r in payload['coverage']))
    return payload

def export(root,payload,render_dir=None):
    config=json.loads((root/'daily_config.json').read_text());output=report_path(root,payload['kind'],dt.date.fromisoformat(payload['start']))
    output.parent.mkdir(parents=True,exist_ok=True);temp=output.with_suffix('.pending.xlsx');job=root/'nifty500_data'/('build_'+payload['kind']+'.json')
    if temp.exists():temp.unlink()
    atomic_json(job,{'payload':payload,'output':str(temp),'render_dir':str(render_dir) if render_dir else None})
    try:
        result=subprocess.run([config['node_executable'],str(root/'daily_runtime/build_nifty500.mjs'),str(job)],cwd=root,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=360,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        if not temp.exists():raise RuntimeError('NSE500 export failed: '+result.stdout[-1000:]+result.stderr[-1000:])
        with zipfile.ZipFile(temp) as z:
            if z.testzip():raise ValueError('Invalid exported workbook')
        temp.replace(output)
    finally:
        if temp.exists():temp.unlink()
    return output

def run(root,kind,starts,config,now,logger,offline=False,render_dir=None):
    state_path=root/('nifty500_'+kind+'_state.json');state=json.loads(state_path.read_text()) if state_path.exists() else {'completed':{}}
    saved=[];pending=[]
    for start in starts:
        key=str(start) if kind=='weekly' else start.strftime('%Y-%m');path=report_path(root,kind,start);receipt=state['completed'].get(key)
        if receipt and path.exists() and hashlib.sha256(path.read_bytes()).hexdigest()==receipt['sha256']:continue
        try:
            payload=prepare(root,kind,start,config,now,logger,offline)
            atomic_json(root/'nifty500_data'/'reports'/f'{kind}_{key}.json',payload)
            path=export(root,payload,render_dir)
            state['completed'][key]={'file':path.relative_to(root/'outputs').as_posix(),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'completed_at':now.isoformat(),'constituent_sha256':payload['constituents']['sha256']}
            atomic_json(state_path,state);queue_report(root,'nifty500:'+kind+':'+key,path,key,'created',logger);saved.append(key)
        except Exception as ex:
            pending.append(key);logger.exception('Nifty 500 report pending: %s',ex)
            if sys.stderr:print(str(ex),file=sys.stderr)
    return {'status':'pending' if pending else 'success','at':now.isoformat(),'saved':saved,'pending':pending,'email_notifications':flush_notifications(root,logger)}

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--kind',choices=['weekly','monthly'],required=True)
    parser.add_argument('--offline',action='store_true');parser.add_argument('--render-dir',type=pathlib.Path)
    args=parser.parse_args();now=dt.datetime.now(IST);file=BASE/'nifty500_config.json';config=json.loads(file.read_text()) if file.exists() else DEFAULT_CONFIG
    logs=BASE/'logs';logs.mkdir(exist_ok=True);logger=logging.getLogger('nifty500-'+args.kind);logger.setLevel(logging.INFO)
    handler=logging.handlers.RotatingFileHandler(logs/('nifty500_'+args.kind+'.log'),maxBytes=2_000_000,backupCount=4,encoding='utf-8');handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'));logger.addHandler(handler)
    import msvcrt
    with (BASE/('nifty500_'+args.kind+'.lock')).open('a+b') as lock:
        if lock.seek(0,2)==0:lock.write(b'0');lock.flush()
        lock.seek(0)
        try:msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
        except OSError:return 0
        try:
            init_cache(BASE);status=run(BASE,args.kind,pending_periods(args.kind,config,now),config,now,logger,args.offline,args.render_dir)
            atomic_json(BASE/('nifty500_'+args.kind+'_last_run_status.json'),status)
            if sys.stdout:print(json.dumps(status,indent=2))
            return 1 if status['pending'] else 0
        finally:lock.seek(0);msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1)

if __name__=='__main__':sys.exit(main())
