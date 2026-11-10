import datetime as dt, json, logging, pathlib, shutil, subprocess, sys, tempfile, unittest, zipfile
from unittest.mock import patch
import update_weekly_ohlc as u
import scheduled_update as s

class SchedulerTests(unittest.TestCase):
    def test_completed_week_boundaries(self):
        expected={'2026-11-10':'2026-11-02','2026-11-13':'2026-11-02','2026-11-14':'2026-11-09','2026-11-15':'2026-11-09','2026-11-16':'2026-11-09','2027-01-02':'2026-12-28'}
        for today,monday in expected.items():
            self.assertEqual(str(u.previous_week(dt.date.fromisoformat(today))[0]),monday)

    def test_due_times(self):
        first=dt.datetime(2026,11,14,9,tzinfo=s.IST)
        self.assertEqual(s.due_runs(first,first-dt.timedelta(seconds=1)),[])
        self.assertEqual(s.due_runs(first,first),[first])
        self.assertEqual(len(s.due_runs(first,dt.datetime(2026,11,30,12,tzinfo=s.IST))),3)

    def test_holidays_and_cache(self):
        holidays,info=u.trading_holidays(2026,offline=True)
        self.assertIn(dt.date(2026,11,10),holidays)
        self.assertNotIn(dt.date(2026,11,11),holidays)
        self.assertEqual(info['year'],2026)

    def test_holiday_closure_is_not_downloaded(self):
        real=u.archive; requested=[]
        def record(day,segment,offline):
            requested.append(day)
            # Reuse cached real data on other days; date field is not relevant here.
            return real(dt.date(2026,11,2)+(day-dt.date(2026,11,9)),segment,True)
        with patch.object(u,'archive',record):
            payload=u.prepare(dt.date(2026,11,9),dt.date(2026,11,13),holidays=[dt.date(2026,11,10)],offline=True)
        self.assertNotIn(dt.date(2026,11,10),requested)
        self.assertEqual(len(payload['daily']),218*4)

    def test_delayed_runs_idempotence_and_retry(self):
        with tempfile.TemporaryDirectory() as folder:
            root=pathlib.Path(folder); (root/'outputs').mkdir()
            config={'first_run':'2026-11-14T09:00:00+05:30'}
            calls=[]; fail={dt.date(2026,11,9)}
            def invoke(monday):
                calls.append(monday)
                if monday in fail: raise OSError('Simulated network outage')
                with zipfile.ZipFile(s.report_path(root,monday),'w') as z: z.writestr('xl/workbook.xml','<workbook/>')
            logger=logging.getLogger('tests'); logger.addHandler(logging.NullHandler())
            now=dt.datetime(2026,11,23,12,tzinfo=s.IST)
            with self.assertRaises(RuntimeError): s.run_pending(root,config,now,invoke,logger)
            self.assertEqual(calls,[dt.date(2026,11,9),dt.date(2026,11,16)])
            state=json.loads((root/'schedule_state.json').read_text())
            self.assertNotIn('2026-11-09',state['completed'])
            self.assertIn('2026-11-16',state['completed'])
            fail.clear(); calls.clear()
            s.run_pending(root,config,now,invoke,logger)
            self.assertEqual(calls,[dt.date(2026,11,9)])
            calls.clear(); s.run_pending(root,config,now,invoke,logger); self.assertEqual(calls,[])

    def test_real_updater_integration(self):
        # Full historical generation using official validated archives, separate output.
        with tempfile.TemporaryDirectory() as folder:
            result=subprocess.run([sys.executable,str(u.BASE/'update_weekly_ohlc.py'),'--week-start','2026-11-02','--as-of','2026-11-07','--offline','--auto-holidays','--output-dir',folder],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            from openpyxl import load_workbook
            workbook=load_workbook(pathlib.Path(folder)/'NSE_FO_Weekly_OHLC_2026-11-02_2026-11-06.xlsx',data_only=True)
            rows=[list(r) for r in workbook['Weekly OHLC'].iter_rows(min_row=2,max_row=219,max_col=5,values_only=True)]
            original=json.loads((u.BASE/'outputs/NSE_FO_Weekly_OHLC_2026-11-02_2026-11-06.json').read_text())
            self.assertEqual([list(r) for r in rows],original['calculated'])
            self.assertEqual(len(rows),218)

if __name__=='__main__': unittest.main(verbosity=2)
