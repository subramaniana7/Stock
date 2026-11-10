"""Plain-text report notifications; durable outbox; secrets in Windows Credential Manager."""
import contextlib, ctypes, datetime as dt, hashlib, json, logging, os, pathlib
import smtplib, ssl, time, uuid
from email.message import EmailMessage

IST=dt.timezone(dt.timedelta(hours=5,minutes=30))
CREDENTIAL_TARGET='NSEStockReports:GmailSMTP'

def atomic_json(path,data):
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_name(path.name+'.'+uuid.uuid4().hex+'.tmp')
    try:
        tmp.write_text(json.dumps(data,indent=2),encoding='utf-8'); tmp.replace(path)
    finally:
        if tmp.exists(): tmp.unlink()

def credential_api():
    from ctypes import wintypes as wt
    class CREDENTIAL(ctypes.Structure):
        _fields_=[('Flags',wt.DWORD),('Type',wt.DWORD),('TargetName',wt.LPWSTR),
                  ('Comment',wt.LPWSTR),('LastWritten',wt.FILETIME),('CredentialBlobSize',wt.DWORD),
                  ('CredentialBlob',ctypes.POINTER(ctypes.c_ubyte)),('Persist',wt.DWORD),
                  ('AttributeCount',wt.DWORD),('Attributes',ctypes.c_void_p),('TargetAlias',wt.LPWSTR),('UserName',wt.LPWSTR)]
    api=ctypes.WinDLL('Advapi32.dll',use_last_error=True)
    api.CredWriteW.argtypes=[ctypes.POINTER(CREDENTIAL),wt.DWORD]; api.CredWriteW.restype=wt.BOOL
    api.CredReadW.argtypes=[wt.LPCWSTR,wt.DWORD,wt.DWORD,ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]; api.CredReadW.restype=wt.BOOL
    api.CredFree.argtypes=[ctypes.c_void_p]; api.CredFree.restype=None
    return api,CREDENTIAL

def store_password(username,password,target=CREDENTIAL_TARGET):
    api,kind=credential_api(); data=password.encode('utf-16-le')
    blob=(ctypes.c_ubyte*len(data)).from_buffer_copy(data)
    cred=kind(); cred.Type=1; cred.TargetName=target; cred.UserName=username
    cred.CredentialBlobSize=len(data);cred.CredentialBlob=blob;cred.Persist=2
    if not api.CredWriteW(ctypes.byref(cred),0): raise ctypes.WinError(ctypes.get_last_error())

def read_password(target=CREDENTIAL_TARGET):
    api,kind=credential_api(); pointer=ctypes.POINTER(kind)()
    if not api.CredReadW(target,1,0,ctypes.byref(pointer)):
        raise RuntimeError('Email sender is not configured. Run setup_email.py under the scheduled-task Windows account.')
    try:
        cred=pointer.contents
        return cred.UserName,ctypes.string_at(cred.CredentialBlob,cred.CredentialBlobSize).decode('utf-16-le')
    finally: api.CredFree(pointer)

@contextlib.contextmanager
def outbox_lock(root):
    folder=root/'email_outbox';folder.mkdir(exist_ok=True)
    with (folder/'outbox.lock').open('a+b') as handle:
        import msvcrt
        if handle.seek(0,2)==0: handle.write(b'0');handle.flush()
        acquired=False
        for _ in range(50):
            try:
                handle.seek(0);msvcrt.locking(handle.fileno(),msvcrt.LK_NBLCK,1);acquired=True;break
            except OSError: time.sleep(.1)
        if not acquired: raise RuntimeError('Notification outbox is busy; retry later')
        try: yield folder
        finally: handle.seek(0);msvcrt.locking(handle.fileno(),msvcrt.LK_UNLCK,1)

def configuration(root):
    path=root/'email_config.json'
    if not path.exists(): return None
    config=json.loads(path.read_text(encoding='utf-8'))
    return config if config.get('enabled') else None

def queue_report(root,event_id,report_path,period,action,logger=None):
    logger=logger or logging.getLogger('nse-email')
    try:
        config=configuration(root)
        if not config: return False
        report_path=pathlib.Path(report_path)
        if not report_path.is_file(): raise ValueError('Report file must exist before email is queued')
        key=hashlib.sha256(event_id.encode()).hexdigest()
        folder=root/'email_outbox';folder.mkdir(exist_ok=True)
        path=folder/(key+'.json')
        if path.exists():return False
        # Publish atomically without waiting for another task's SMTP connection.
        # A hard link creates the final name only if it does not already exist.
        temp=folder/(key+'.'+uuid.uuid4().hex+'.tmp')
        try:
            temp.write_text(json.dumps({'id':event_id,'key':key,'filename':report_path.name,'period':period,
                'action':action,'saved_at':dt.datetime.now(IST).isoformat(),'status':'pending','attempts':0,
                'recipient':config['recipient'],'sender':config['sender']},indent=2),encoding='utf-8')
            try:os.link(temp,path)
            except FileExistsError:return False
        finally:
            if temp.exists():temp.unlink()
        return True
    except Exception as ex:
        logger.warning('Could not queue report email (%s); report remains saved.',type(ex).__name__)
        return False

def make_message(event):
    message=EmailMessage()
    message['From']=event['sender'];message['To']=event['recipient']
    message['Subject']='NSE Excel report '+event['action']+': '+event['filename']
    message['Message-ID']='<nse-'+event['key']+'@stock-report.local>'
    message.set_content('Your NSE Excel report was saved successfully.\n\n'
        +'File: '+event['filename']+'\n'
        +'Report period / trading date: '+event['period']+'\n'
        +'Action: '+event['action']+'\n'
        +'Saved at (India time): '+event['saved_at']+'\n\n'
        +'The file is available in your local NSE Weekly OHLC outputs folder.\n')
    return message

def smtp_send(config,event,credential_reader=read_password):
    username,password=credential_reader(config.get('credential_target',CREDENTIAL_TARGET))
    if username.lower()!=config['sender'].lower(): raise ValueError('Stored sender does not match email configuration')
    with smtplib.SMTP_SSL(config['smtp_host'],int(config.get('smtp_port',465)),
            context=ssl.create_default_context(),timeout=30) as server:
        server.login(username,password)
        server.send_message(make_message(event),from_addr=event['sender'],to_addrs=[event['recipient']])

def flush_notifications(root,logger=None,send=None,now=None):
    logger=logger or logging.getLogger('nse-email');now=now or dt.datetime.now(IST)
    result={'sent':0,'pending':0}
    try:
        config=configuration(root)
        if not config:return result
        sender=send or smtp_send
        with outbox_lock(root) as folder:
            for path in sorted(folder.glob('*.json')):
                event=json.loads(path.read_text(encoding='utf-8'))
                if event['status']=='sent':continue
                due=event.get('next_attempt')
                if due and dt.datetime.fromisoformat(due)>now:result['pending']+=1;continue
                try:
                    # No report is attached or read by this sender.
                    sender(config,event)
                    event.update(status='sent',sent_at=now.isoformat());event.pop('last_error',None)
                    result['sent']+=1
                    logger.info('Email sent: %s (%s)',event['filename'],event['period'])
                except Exception as ex:
                    event['attempts']+=1;event['last_error']=type(ex).__name__
                    event['next_attempt']=(now+dt.timedelta(minutes=15)).isoformat()
                    result['pending']+=1
                    logger.warning('Email pending: %s (%s); will retry.',event['filename'],type(ex).__name__)
                atomic_json(path,event)
    except Exception as ex:
        logger.warning('Email delivery deferred (%s); report generation is unaffected.',type(ex).__name__)
        result['pending']+=1
    return result
