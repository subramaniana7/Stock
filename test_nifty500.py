import datetime as dt,unittest
import nifty500_update as n

class Nifty500Tests(unittest.TestCase):
    def test_dates_names_and_rollover(self):
        root=n.pathlib.Path('reports');self.assertEqual(n.report_path(root,'weekly',dt.date(2026,11,9)).name,'NSE500 Weekly Nov 9-13 2026.xlsx')
        self.assertEqual(n.report_path(root,'monthly',dt.date(2026,10,1)).name,'NSE500 Monthly Oct 2026.xlsx')
        self.assertEqual(n.pending_periods('weekly',n.DEFAULT_CONFIG,dt.datetime(2026,11,14,8,59,tzinfo=n.IST)),[])
        self.assertEqual(n.pending_periods('weekly',n.DEFAULT_CONFIG,dt.datetime(2026,11,14,9,tzinfo=n.IST)),[dt.date(2026,11,9)])
        self.assertEqual(n.report_path(root,'weekly',dt.date(2026,11,30)).parent.name,'2026-11')
    def test_series_isin_matching_and_dummy(self):
        member={'Symbol':'NEW','ISIN Code':'TESTISIN'}
        row={'SsnId':'F1','SctySrs':'BE','ISIN':'TESTISIN','TckrSymb':'OLD','OpnPric':'10','HghPric':'12','LwPric':'9','ClsPric':'11'}
        self.assertEqual(n.select_cash(member,[row]),[10,12,9,11,'OLD','BE'])
        row['SctySrs']='RR';self.assertEqual(n.select_cash(member,[row])[-1],'RR')
        row['SctySrs']='BL';self.assertIsNone(n.select_cash(member,[row]))
        self.assertIsNone(n.select_cash({'Symbol':'DUMMY','ISIN Code':'DUMMYISIN'},[]))
    def test_all_members_retained_missing_marked_unavailable(self):
        day=dt.date(2026,10,1);members=[{'Symbol':'A','ISIN Code':'TESTISIN'},{'Symbol':'DUMMY','ISIN Code':'DUMMYISIN'}]
        row={'SsnId':'F1','SctySrs':'EQ','ISIN':'TESTISIN','TckrSymb':'A','OpnPric':'10','HghPric':'12','LwPric':'9','ClsPric':'11'}
        info={'source':'Official list','downloaded_at':'2026-11-14T09:00:00+05:30','sha256':'test'}
        p=n.aggregate('monthly',day,[day],members,info,{day:[row]},[],{'source':'NSE calendar'})
        self.assertEqual(p['count'],2);self.assertEqual(p['calculated'],[['A',10,12,9,11],['DUMMY',None,None,None,None]])
        self.assertEqual(p['coverage'][1][1],'Unavailable: no cash prices in period')
    def test_duplicate_or_invalid_list_rejected(self):
        with self.assertRaises(ValueError):n.parse_constituents(b'<html>Blocked</html>')
        with self.assertRaises(ValueError):n.parse_constituents(b'Company Name,Industry,Symbol,Series,ISIN Code\nA,Industry,A,EQ,INE000000001\nA,Industry,A,EQ,INE000000001\n')

if __name__=='__main__':unittest.main(verbosity=2)
