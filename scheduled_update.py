"""Unattended weekly runner. Catch up every scheduled week since installation."""
import argparse, contextlib, datetime as dt, hashlib, json, logging, logging.handlers
import os, pathlib, subprocess, sys, zipfile
from email_notifications import flush_notifications

BASE=pathlib.Path(__file__).resolve().parent
IST=dt.timezone(dt.timedelta(hours=5,minutes=30))

def atomic_json(path,data):
    temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps(data,indent=2),encoding='utf-8')
    temp.replace(path)

def due_runs(first_run, now):
    due=[]; cursor=first_run
    while cursor<=now:
        due.append(cursor)
        cursor+=dt.timedelta(days=7)
    return due

def report_path(root, monday):
    return root/'outputs'/f'NSE_FO_Weekly_OHLC_{monday}_{monday+dt.timedelta(days=4)}.xlsx'

def intact_receipt(root,monday,receipt):
    path=report_path(root,monday)
    return bool(receipt and path.is_file() and hashlib.sha256(path.read_bytes()).hexdigest()==receipt.get('sha256'))

def run_pending(root,config,now,invoke,logger):
    first=dt.datetime.fromisoformat(config['first_run'])
    if first.weekday()!=5 or (first.hour,first.minute)!=(9,0) or first.utcoffset()!=dt.timedelta(hours=5,minutes=30):
        raise ValueError('Schedule must start on Saturday at 09:00 IST')
    state_path=root/'schedule_state.json'
    state=json.loads(state_path.read_text(encoding='utf-8')) if state_path.exists() else {'completed':{}}
    failures=[]
    for scheduled in due_runs(first,now):
        monday=scheduled.date()-dt.timedelta(days=5); key=str(monday)
        receipt=state['completed'].get(key)
        if intact_receipt(root,monday,receipt): continue
        logger.info('Starting report for %s to %s; scheduled %s',monday,monday+dt.timedelta(days=4),scheduled.isoformat())
        try:
            invoke(monday)
            path=report_path(root,monday)
            with zipfile.ZipFile(path) as z:
                if z.testzip() or 'xl/workbook.xml' not in z.namelist(): raise ValueError('Invalid exported workbook')
            state['completed'][key]={'scheduled_for':scheduled.isoformat(),'completed_at':now.isoformat(),'file':path.name,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()}
            atomic_json(state_path,state)
            logger.info('SUCCESS: %s',path)
        except Exception as ex:
            failures.append(key); logger.exception('FAILED report for %s: %s',key,ex)
    if failures: raise RuntimeError('Failed weeks remain pending for retry: '+', '.join(failures))
    logger.info('Finished: all reports due as of %s are complete (%s scheduled weeks).',now.isoformat(),len(due_runs(first,now)))

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check-only',action='store_true')
    args=parser.parse_args()
    config=json.loads((BASE/'schedule_config.json').read_text(encoding='utf-8'))
    now=dt.datetime.now(IST)
    if args.check_only:
        print(json.dumps({'now':now.isoformat(),'first_run':config['first_run'],'due_runs':[x.isoformat() for x in due_runs(dt.datetime.fromisoformat(config['first_run']),now)]},indent=2)); return 0
    logs=BASE/'logs'; logs.mkdir(exist_ok=True)
    logger=logging.getLogger('nse-weekly'); logger.setLevel(logging.INFO)
    handler=logging.handlers.RotatingFileHandler(logs/'weekly_update.log',maxBytes=2_000_000,backupCount=4,encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s')); logger.addHandler(handler)
    logger.info('Runner started; India time %s; Python %s',now.isoformat(),sys.executable)
    # Windows task ignores overlapping instances; lock also protects manual runs.
    import msvcrt
    lock=(BASE/'scheduled_update.lock').open('a+b')
    try:
        lock.seek(0); lock.write(b'0'); lock.flush(); lock.seek(0)
        try: msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
        except OSError:
            logger.info('Another weekly runner is active; exiting.'); return 0
        def invoke(monday):
            python=pathlib.Path(sys.executable).with_name('python.exe')
            command=[str(python),str(BASE/'update_weekly_ohlc.py'),'--week-start',str(monday),'--auto-holidays']
            result=subprocess.run(command,cwd=BASE,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=1200,creationflags=subprocess.CREATE_NO_WINDOW)
            if result.stdout: logger.info('%s',result.stdout.strip())
            if result.stderr: logger.warning('%s',result.stderr.strip())
            if result.returncode: raise RuntimeError(f'Updater exited with code {result.returncode}')
        try:
            run_pending(BASE,config,now,invoke,logger)
            flush_notifications(BASE,logger)
            atomic_json(BASE/'last_run_status.json',{'status':'success','at':now.isoformat(),'message':'All due weekly reports complete'})
            return 0
        except Exception as ex:
            logger.exception('Scheduled run failed')
            atomic_json(BASE/'last_run_status.json',{'status':'failed','at':now.isoformat(),'error':str(ex)})
            return 1
        finally:
            lock.seek(0); msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1)
    finally:
        lock.close()

if __name__=='__main__':
    sys.exit(main())
