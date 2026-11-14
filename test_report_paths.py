import datetime as dt,hashlib,json,pathlib,tempfile,unittest
from report_paths import weekly_path,monthly_path,daily_path
from migrate_output_folders import migrate

class PathTests(unittest.TestCase):
    def test_names_month_and_year_boundaries(self):
        root=pathlib.Path('reports')
        self.assertEqual(str(weekly_path(root,dt.date(2026,11,9)).relative_to(root)).replace('\\','/'),'outputs/2026-11/Weekly Nov 9-13 2026.xlsx')
        self.assertEqual(weekly_path(root,dt.date(2026,11,30)).name,'Weekly Nov 30-Dec 4 2026.xlsx')
        self.assertEqual(weekly_path(root,dt.date(2026,12,28)).name,'Weekly Dec 28 2026-Jan 1 2027.xlsx')
        self.assertEqual(monthly_path(root,dt.date(2026,10,1)).name,'Monthly Oct 2026.xlsx')
        self.assertNotEqual(monthly_path(root,dt.date(2026,10,1)),monthly_path(root,dt.date(2027,10,1)))
        self.assertEqual(daily_path(root,dt.date(2026,11,30)).parent.name,'2026-11')
    def test_migration_hashes_receipts_no_duplicate_notifications(self):
        with tempfile.TemporaryDirectory() as folder:
            root=pathlib.Path(folder);output=root/'outputs';output.mkdir()
            old=output/'NSE_FO_Weekly_OHLC_2026-11-09_2026-11-13.xlsx';old.write_bytes(b'unchanged workbook')
            (root/'schedule_state.json').write_text(json.dumps({'completed':{'2026-11-09':{'file':old.name,'sha256':hashlib.sha256(old.read_bytes()).hexdigest()}}}))
            (root/'email_outbox').mkdir();sent={'status':'sent','filename':old.name,'id':'weekly:2026-11-09'}
            (root/'email_outbox/sent.json').write_text(json.dumps(sent))
            result=migrate(root);target=weekly_path(root,dt.date(2026,11,9))
            self.assertEqual(result['moved'],1);self.assertFalse(old.exists());self.assertEqual(target.read_bytes(),b'unchanged workbook')
            receipt=json.loads((root/'schedule_state.json').read_text())['completed']['2026-11-09']
            self.assertEqual(receipt['file'],'2026-11/Weekly Nov 9-13 2026.xlsx')
            self.assertEqual(json.loads((root/'email_outbox/sent.json').read_text()),sent)
            self.assertEqual(migrate(root)['moved'],0)
    def test_collision_preserves_both_files(self):
        with tempfile.TemporaryDirectory() as folder:
            root=pathlib.Path(folder);(root/'outputs').mkdir()
            old=root/'outputs/NSE_FO_Monthly_OHLC_2026-10.xlsx';old.write_bytes(b'old')
            target=monthly_path(root,dt.date(2026,10,1));target.parent.mkdir();target.write_bytes(b'existing')
            with self.assertRaises(FileExistsError):migrate(root)
            self.assertEqual(old.read_bytes(),b'old');self.assertEqual(target.read_bytes(),b'existing')

if __name__=='__main__':unittest.main(verbosity=2)
