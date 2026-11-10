"""Local Windows masked credential prompt; secrets never appear in command-line arguments."""
import ctypes,datetime as dt,json,pathlib,smtplib,ssl
from ctypes import wintypes as wt
from email.message import EmailMessage
from email_notifications import store_password,atomic_json,CREDENTIAL_TARGET,flush_notifications

BASE=pathlib.Path(__file__).resolve().parent

def message(text,kind=0x40):
    return ctypes.windll.user32.MessageBoxW(None,text,'NSE report email setup',kind)

def prompt(sender):
    class INFO(ctypes.Structure):
        _fields_=[('cbSize',wt.DWORD),('hwndParent',wt.HWND),('pszMessageText',wt.LPCWSTR),
                  ('pszCaptionText',wt.LPCWSTR),('hbmBanner',wt.HANDLE)]
    api=ctypes.WinDLL('Credui.dll',use_last_error=True)
    api.CredUIPromptForCredentialsW.argtypes=[ctypes.POINTER(INFO),wt.LPCWSTR,wt.HANDLE,wt.DWORD,
        wt.LPWSTR,wt.ULONG,wt.LPWSTR,wt.ULONG,ctypes.POINTER(wt.BOOL),wt.DWORD]
    api.CredUIPromptForCredentialsW.restype=wt.DWORD
    info=INFO();info.cbSize=ctypes.sizeof(INFO)
    info.pszCaptionText='NSE report email setup'
    info.pszMessageText=('Enter the 16-character Gmail APP PASSWORD for '+sender+'.\n'
        'Use the App Password created for NSE Reports, not your regular Google password.\n'
        'Click OK to save securely and send a test email. No attachments.')
    user=ctypes.create_unicode_buffer(sender,514);password=ctypes.create_unicode_buffer(256);save=wt.BOOL(False)
    code=api.CredUIPromptForCredentialsW(ctypes.byref(info),CREDENTIAL_TARGET,None,0,
        user,len(user),password,len(password),ctypes.byref(save),0x40000|0x80|0x2)
    if code==1223:return None
    if code:raise ctypes.WinError(code)
    value=password.value.replace(' ','').strip()
    ctypes.memset(ctypes.addressof(password),0,ctypes.sizeof(password))
    if user.value.lower()!=sender.lower():raise ValueError('The username must match the configured Gmail sender')
    return value

def configure(config,value):
    mail=EmailMessage();mail['From']=config['sender'];mail['To']=config['recipient']
    mail['Subject']='NSE report email notifications are ready'
    mail.set_content('Email notifications for your daily and weekly NSE Excel reports are now configured.\n'
        'You will receive a short message after a workbook is created or a new daily sheet is saved.\n'
        'No Excel files are attached.\n')
    with smtplib.SMTP_SSL(config['smtp_host'],int(config['smtp_port']),context=ssl.create_default_context(),timeout=30) as server:
        server.login(config['sender'],value)
        server.send_message(mail,from_addr=config['sender'],to_addrs=[config['recipient']])
    store_password(config['sender'],value,config.get('credential_target',CREDENTIAL_TARGET))
    atomic_json(BASE/'email_setup_status.json',{'status':'ready','at':dt.datetime.now().astimezone().isoformat(),'test_email_sent':True})
    flush_notifications(BASE)

def main():
    config=json.loads((BASE/'email_config.json').read_text(encoding='utf-8'))
    while True:
        try:
            value=prompt(config['sender'])
            if value is None:
                atomic_json(BASE/'email_setup_status.json',{'status':'cancelled'});return
            if len(value)!=16 or not value.isalnum():
                value='';message('Enter the 16-character Google App Password, not your regular password.',0x30);continue
            configure(config,value);value=''
            message('Test email sent successfully. Daily and weekly report notifications are enabled.');return
        except Exception as ex:
            value=''
            atomic_json(BASE/'email_setup_status.json',{'status':'pending','error':type(ex).__name__})
            if message('Setup did not finish ('+type(ex).__name__+'). Check the App Password and internet, then choose Retry.',0x35)!=4:return

if __name__=='__main__':main()
