import datetime as dt,json,logging,pathlib,tempfile,unittest
from unittest.mock import patch,MagicMock
import email_notifications as e

class EmailTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=pathlib.Path(self.temp.name)
        self.config={'enabled':True,'sender':'sender@gmail.com','recipient':'recipient@gmail.com','smtp_host':'smtp.gmail.com','smtp_port':465}
        e.atomic_json(self.root/'email_config.json',self.config)
        self.report=self.root/'report.xlsx';self.report.write_bytes(b'fake report for queue tests')
        self.logger=logging.getLogger('email-tests');self.logger.addHandler(logging.NullHandler())
    def tearDown(self):self.temp.cleanup()
    def queue(self):return e.queue_report(self.root,'daily:2026-11-09',self.report,'2026-11-09','created',self.logger)
    def test_plain_text_no_attachment_and_dedup(self):
        self.assertTrue(self.queue());self.assertFalse(self.queue())
        sent=[];result=e.flush_notifications(self.root,self.logger,send=lambda c,v:sent.append(v))
        self.assertEqual(result,{'sent':1,'pending':0});self.assertEqual(len(sent),1)
        message=e.make_message(sent[0]);self.assertEqual(message['To'],'recipient@gmail.com')
        self.assertFalse(message.is_multipart());self.assertEqual(list(message.iter_attachments()),[])
        self.assertIn('report.xlsx',message.get_content());self.assertNotIn('fake report',message.get_content())
        self.assertFalse(self.queue());e.flush_notifications(self.root,self.logger,send=lambda c,v:self.fail('duplicate email'))
    def test_failure_retried_without_recreating_report(self):
        self.queue();before=self.report.read_bytes();now=dt.datetime(2026,11,10,18,tzinfo=e.IST)
        failed=lambda c,v: (_ for _ in ()).throw(OSError('Simulated connection failure'))
        result=e.flush_notifications(self.root,self.logger,send=failed,now=now)
        self.assertEqual(result,{'sent':0,'pending':1});self.assertEqual(before,self.report.read_bytes())
        result=e.flush_notifications(self.root,self.logger,send=lambda c,v:self.fail('too early'),now=now+dt.timedelta(minutes=14))
        self.assertEqual(result['pending'],1)
        sent=[];result=e.flush_notifications(self.root,self.logger,send=lambda c,v:sent.append(v),now=now+dt.timedelta(minutes=15))
        self.assertEqual(result['sent'],1);self.assertEqual(len(sent),1)
    def test_no_event_for_missing_file(self):
        self.report.unlink();self.assertFalse(self.queue());self.assertFalse(list((self.root/'email_outbox').glob('*.json')) if (self.root/'email_outbox').exists() else [])
    def test_unconfigured_email_does_not_fail_report(self):
        (self.root/'email_config.json').unlink();self.assertFalse(self.queue());self.assertEqual(e.flush_notifications(self.root),{'sent':0,'pending':0})
    def test_smtp_secure_sender_and_recipient(self):
        self.queue();event=json.loads(next((self.root/'email_outbox').glob('*.json')).read_text())
        server=MagicMock()
        with patch.object(e.smtplib,'SMTP_SSL') as smtp:
            smtp.return_value.__enter__.return_value=server
            e.smtp_send(self.config,event,lambda target:('sender@gmail.com','mock-app-password'))
            server.login.assert_called_once_with('sender@gmail.com','mock-app-password')
            self.assertEqual(server.send_message.call_args.kwargs['to_addrs'],['recipient@gmail.com'])
            message=server.send_message.call_args.args[0];self.assertEqual(list(message.iter_attachments()),[])
    def test_sender_mismatch_blocks_delivery(self):
        self.queue();event=json.loads(next((self.root/'email_outbox').glob('*.json')).read_text())
        with self.assertRaises(ValueError): e.smtp_send(self.config,event,lambda target:('different@gmail.com','unused'))
    def test_queue_while_another_task_is_sending(self):
        with e.outbox_lock(self.root):
            self.assertTrue(self.queue())
        self.assertEqual(len(list((self.root/'email_outbox').glob('*.json'))),1)

if __name__=='__main__':unittest.main(verbosity=2)
