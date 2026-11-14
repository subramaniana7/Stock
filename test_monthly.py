import datetime as dt,logging,pathlib,tempfile,unittest
from unittest.mock import patch
import monthly_update as m

class MonthlyTests(unittest.TestCase):
    def test_month_year_rollover_and_leap_year(self):
        self.assertEqual(m.next_month(dt.date(2026,12,1)),dt.date(2027,1,1))
        self.assertEqual(m.previous_month(dt.date(2027,1,1)),dt.date(2026,12,1))
        self.assertEqual(m.next_month(dt.date(2028,2,1))-dt.timedelta(days=1),dt.date(2028,2,29))
    def test_first_day_9am_due_and_multi_month_catchup(self):
        self.assertEqual(m.due_months('2026-10',dt.datetime(2026,11,1,8,59,tzinfo=m.IST)),[])
        self.assertEqual(m.due_months('2026-10',dt.datetime(2026,11,1,9,tzinfo=m.IST)),[dt.date(2026,10,1)])
        self.assertEqual(m.due_months('2026-10',dt.datetime(2027,1,10,12,tzinfo=m.IST)),[dt.date(2026,10,1),dt.date(2026,11,1),dt.date(2026,12,1)])
    def test_october_closures_and_special_sessions(self):
        days=m.session_dates(dt.date(2026,10,1),{dt.date(2026,10,2),dt.date(2026,10,20)})
        self.assertEqual(len(days),20);self.assertEqual(days[0],dt.date(2026,10,1));self.assertEqual(days[-1],dt.date(2026,10,30))
        special=dt.date(2026,11,8)
        self.assertIn(special,m.session_dates(dt.date(2026,11,1),{special},[special]))
    def test_monthly_aggregation_and_eligibility_union(self):
        days=[dt.date(2026,10,d) for d in [1,5,30]]
        def row(o,h,l,c):return dict(zip(['OpnPric','HghPric','LwPric','ClsPric'],[o,h,l,c]))
        cash={days[0]:{'A':row(10,13,9,12),'B':row(20,22,19,21)},days[1]:{'A':row(12,18,11,15),'B':row(21,30,18,25)},days[2]:{'A':row(15,17,8,16),'B':row(25,28,20,26)}}
        universe={days[0]:{'A'},days[1]:{'A','B'},days[2]:{'B'}}
        p=m.aggregate(days[0],days,cash,universe,[],{'source':'official NSE calendar'})
        self.assertEqual(p['calculated'],[['A',10,18,8,16],['B',20,30,18,26]])
        self.assertEqual(len(p['daily']),6)
        cash[days[0]].pop('B')
        with self.assertRaises(ValueError):m.aggregate(days[0],days,cash,universe,[],{'source':'NSE'})
    def test_failed_month_remains_pending(self):
        with tempfile.TemporaryDirectory() as folder:
            root=pathlib.Path(folder);logger=logging.getLogger('test-monthly');logger.addHandler(logging.NullHandler())
            with patch.object(m,'prepare_month',side_effect=OSError('Missing archive')):
                status=m.run_months(root,[dt.date(2026,10,1)],{},dt.datetime(2026,11,14,tzinfo=m.IST),logger)
            self.assertEqual(status['pending_months'],['2026-10']);self.assertFalse((root/'monthly_state.json').exists())
    def test_successful_month_is_not_recreated(self):
        with tempfile.TemporaryDirectory() as folder:
            root=pathlib.Path(folder);(root/'outputs').mkdir();start=dt.date(2026,10,1)
            path=m.report_path(root,start);path.write_bytes(b'previously verified monthly workbook')
            m.atomic_json(root/'monthly_state.json',{'completed':{'2026-10':{'sha256':m.hashlib.sha256(path.read_bytes()).hexdigest()}}})
            with patch.object(m,'prepare_month',side_effect=AssertionError('Completed month downloaded again')):
                result=m.run_months(root,[start],{},dt.datetime(2026,11,14,tzinfo=m.IST),logging.getLogger('monthly-test'))
            self.assertEqual(result['saved_months'],[]);self.assertEqual(result['pending_months'],[])
    def test_isin_confirmed_symbol_rename(self):
        days=[dt.date(2026,10,1),dt.date(2026,10,30)]
        def row(o,h,l,c):return dict(zip(['OpnPric','HghPric','LwPric','ClsPric','ISIN'],[o,h,l,c,'TESTISIN']))
        cash={days[0]:{'OLD':row(10,15,9,12)},days[1]:{'NEW':row(12,18,11,16)}}
        payload=m.aggregate(days[0],days,cash,{days[0]:{'OLD'},days[1]:{'NEW'}},[],{'source':'NSE'})
        self.assertEqual(payload['count'],1);self.assertEqual(payload['calculated'],[['NEW',10,18,9,16]])
        self.assertEqual([r[-1] for r in payload['daily']],['OLD','NEW'])
        self.assertEqual(len(payload['symbol_mappings']),1)

if __name__=='__main__':unittest.main(verbosity=2)
