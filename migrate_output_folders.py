"""Move existing reports without rewriting workbooks or resending notifications."""
import argparse,datetime as dt,hashlib,json,pathlib,re,shutil
from report_paths import weekly_path,monthly_path,daily_path

def plan(root):
    output=root/'outputs';moves=[]
    for source in sorted(output.iterdir()):
        if not source.is_file():continue
        weekly=re.match(r'NSE_FO_Weekly_OHLC_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})(.*)$',source.name)
        monthly=re.match(r'NSE_FO_Monthly_OHLC_(\d{4}-\d{2})(.*)$',source.name)
        daily=re.match(r'NSE_FO_Daily_OHLC_(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})(.*)$',source.name)
        if weekly:
            new=weekly_path(root,dt.date.fromisoformat(weekly[1]));target=new.with_name(new.stem+weekly[3])
        elif monthly:
            new=monthly_path(root,dt.date.fromisoformat(monthly[1]+'-01'));target=new.with_name(new.stem+monthly[2])
        elif daily:
            new=daily_path(root,dt.date.fromisoformat(daily[1]));target=new.with_name(source.name)
        else:continue
        source=source.resolve();target=target.resolve()
        if output.resolve() not in source.parents or output.resolve() not in target.parents:raise ValueError('Report path is outside outputs')
        moves.append((source,target))
    targets=[str(t) for _,t in moves]
    if len(set(targets))!=len(targets):raise ValueError('Multiple source files map to the same destination')
    for _,target in moves:
        if target.exists():raise FileExistsError(f'Preserving existing destination: {target}')
    return moves

def atomic_json(path,value):
    tmp=path.with_suffix('.migration.tmp');tmp.write_text(json.dumps(value,indent=2),encoding='utf-8');tmp.replace(path)

def migrate(root):
    root=root.resolve();moves=plan(root)
    if not moves:return {'moved':0,'message':'No legacy root-level report files to migrate'}
    backup=root/'backups'/('before_month_folders_'+dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f'));backup.mkdir(parents=True)
    state_files=[root/'schedule_state.json',root/'monthly_state.json']
    for path in state_files:
        if path.exists():shutil.copy2(path,backup/path.name)
    manifest=[]
    for source,target in moves:
        manifest.append({'from':str(source),'to':str(target),'sha256':hashlib.sha256(source.read_bytes()).hexdigest()})
    atomic_json(backup/'moves.json',manifest)
    done=[]
    try:
        for source,target in moves:
            target.parent.mkdir(parents=True,exist_ok=True);source.rename(target);done.append((source,target))
        for item in manifest:
            if hashlib.sha256(pathlib.Path(item['to']).read_bytes()).hexdigest()!=item['sha256']:raise ValueError('Report contents changed during migration')
    except Exception:
        for source,target in reversed(done):
            if target.exists() and not source.exists():target.rename(source)
        raise
    for path in state_files:
        if not path.exists():continue
        state=json.loads(path.read_text(encoding='utf-8'))
        for key,receipt in state.get('completed',{}).items():
            target=weekly_path(root,dt.date.fromisoformat(key)) if path.name=='schedule_state.json' else monthly_path(root,dt.date.fromisoformat(key+'-01'))
            if target.exists():
                receipt['file']=target.relative_to(root/'outputs').as_posix()
                receipt['sha256']=hashlib.sha256(target.read_bytes()).hexdigest()
        atomic_json(path,state)
    # Historical sent mail remains immutable. Only unsent events receive the new location.
    for path in (root/'email_outbox').glob('*.json'):
        event=json.loads(path.read_text(encoding='utf-8'))
        if event.get('status')=='sent':continue
        for source,target in moves:
            if source.name==event.get('filename'):
                shutil.copy2(path,backup/('email_'+path.name));event['filename']=target.name
                event['relative_file']=target.relative_to(root/'outputs').as_posix();atomic_json(path,event);break
    return {'moved':len(moves),'backup':str(backup),'files':[str(t.relative_to(root/'outputs')) for _,t in moves]}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('folder',type=pathlib.Path);parser.add_argument('--apply',action='store_true');args=parser.parse_args()
    print(json.dumps(migrate(args.folder) if args.apply else [{'from':str(s),'to':str(t)} for s,t in plan(args.folder)],indent=2))
