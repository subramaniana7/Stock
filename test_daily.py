import datetime as dt, json, logging, pathlib, tempfile, unittest, zipfile
from unittest.mock import patch
import daily_update as d
from openpyxl import load_workbook

class DailyTests(unittest.TestCase):
    def setUp(self):
        self.logger=logging.getLogger('daily-tests'); self.logger.addHandler(logging.NullHandler())
        self.config={'first_date':'2026-11-09'}
    def test_cutoff_holidays_weekends(self):
        first=dt.date(2026,11,9); holiday={dt.date(2026,11,10)}
        self.assertEqual(d.due_dates(first,dt.datetime(2026,11,9,16,29,tzinfo=d.IST),holiday),[])
        self.assertEqual(d.due_dates(first,dt.datetime(2026,11,9,16,30,tzinfo=d.IST),holiday),[first])
        due=d.due_dates(first,dt.datetime(2026,11,15,20,tzinfo=d.IST),holiday)
        self.assertEqual(due,[dt.date(2026,11,i) for i in [9,11,12,13]])
    def test_week_rollover(self):
        self.assertEqual(d.monday_of(dt.date(2026,11,13)),dt.date(2026,11,9))
        self.assertEqual(d.monday_of(dt.date(2026,11,16)),dt.date(2026,11,16))
        self.assertNotEqual(d.workbook_path(pathlib.Path('.'),dt.date(2026,11,9)),d.workbook_path(pathlib.Path('.'),dt.date(2026,11,16)))
    def test_catchup_preserves_days_and_no_duplicates(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=pathlib.Path(tmp); calls=[]; builds=[]
            def collect(day,offline): calls.append(day); return {'date':str(day),'rows':[['TEST',10,12,9,11]]}
            def build(root,config,payload,render):
                builds.append(sorted(payload['days'])); out=d.workbook_path(root,dt.date.fromisoformat(payload['week_start'])); out.parent.mkdir(exist_ok=True)
                # Minimal workbook metadata sufficient for the resume validator.
                names=[d.sheet_name(dt.date.fromisoformat(x)) for x in sorted(payload['days'])]+['Sources','CPR Guide']
                with zipfile.ZipFile(out,'w') as z:
                    z.writestr('xl/workbook.xml','<workbook xmlns="'+d.NS['s']+'"><sheets>'+''.join(f'<sheet name="{name}"/>' for name in names)+'</sheets></workbook>')
                return out
            with patch.object(d,'load_calendar',return_value=({dt.date(2026,11,10)},{'year':2026})):
                result=d.run_updates(root,self.config,dt.datetime(2026,11,11,20,tzinfo=d.IST),self.logger,collector=collect,builder=build)
                self.assertEqual(result['added_dates'],['2026-11-09','2026-11-11'])
                result=d.run_updates(root,self.config,dt.datetime(2026,11,12,20,tzinfo=d.IST),self.logger,collector=collect,builder=build)
                self.assertEqual(result['added_dates'],['2026-11-12'])
                self.assertEqual(builds[-1],['2026-11-09','2026-11-11','2026-11-12'])
                count=len(builds); d.run_updates(root,self.config,dt.datetime(2026,11,12,21,tzinfo=d.IST),self.logger,collector=collect,builder=build)
                self.assertEqual(len(builds),count)
                d.run_updates(root,self.config,dt.datetime(2026,11,16,20,tzinfo=d.IST),self.logger,collector=collect,builder=build)
                self.assertTrue(d.workbook_path(root,dt.date(2026,11,16)).exists())
                self.assertTrue(d.workbook_path(root,dt.date(2026,11,9)).exists())
    def test_locked_workbook_does_not_commit(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=pathlib.Path(tmp)
            with patch.object(d,'load_calendar',return_value=(set(),{})):
                result=d.run_updates(root,self.config,dt.datetime(2026,11,9,20,tzinfo=d.IST),self.logger,
                    collector=lambda day,offline:{'date':str(day),'rows':[]},builder=lambda *a: (_ for _ in ()).throw(PermissionError('Excel has file open')))
                self.assertEqual(result['pending_dates'],['2026-11-09'])
                self.assertFalse((root/'daily_data/weeks/2026-11-09.json').exists())
    def test_missing_source_stays_pending(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=pathlib.Path(tmp)
            with patch.object(d,'load_calendar',return_value=(set(),{})):
                result=d.run_updates(root,self.config,dt.datetime(2026,11,9,20,tzinfo=d.IST),self.logger,
                    collector=lambda *a: (_ for _ in ()).throw(OSError('File not published')))
                self.assertEqual(result['pending_dates'],['2026-11-09'])
                self.assertFalse((root/'outputs').exists())
    def test_cached_official_daily_values(self):
        d.init_sources(d.BASE)
        day=d.collect_day(dt.date(2026,11,2),offline=True)
        self.assertEqual(len(day['rows']),218)
        self.assertEqual([r for r in day['rows'] if r[0]=='RELIANCE'],[['RELIANCE',1228.1,1300.0,1214.2,1300.0]])

if __name__=='__main__': unittest.main(verbosity=2)
