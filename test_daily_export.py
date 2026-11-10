"""Real artifact export: append day, locked workbook retry, week rollover, all prices."""
import datetime as dt, hashlib, json, logging, pathlib, tempfile
from unittest.mock import patch
from openpyxl import load_workbook
import daily_update as d

d.init_sources(d.BASE)
config=json.loads((d.BASE/'daily_config.json').read_text())
config.update(first_date='2026-11-02',builder_script=str(d.BASE/'daily_runtime/build_daily.mjs'))
logger=logging.getLogger('export-test'); logger.addHandler(logging.NullHandler())
with tempfile.TemporaryDirectory(prefix='nse-daily-test-') as folder:
    root=pathlib.Path(folder)
    with patch.object(d,'load_calendar',side_effect=lambda root,year,now,offline:d.nse.trading_holidays(year,True)):
        def update(day):
            status=d.run_updates(root,config,dt.datetime(2026,11,day,20,tzinfo=d.IST),logger,offline=True)
            assert status['status']=='success',status
        update(2); update(3)
        first=d.workbook_path(root,dt.date(2026,11,2))
        w=load_workbook(first,data_only=True)
        assert w.sheetnames==['02-Nov-2026','03-Nov-2026','Sources','CPR Guide'],w.sheetnames
        before=hashlib.sha256(first.read_bytes()).hexdigest()
        update(3); assert hashlib.sha256(first.read_bytes()).hexdigest()==before
        with first.open('rb'):
            status=d.run_updates(root,config,dt.datetime(2026,11,4,20,tzinfo=d.IST),logger,offline=True)
            assert status['status']=='pending',status
            assert hashlib.sha256(first.read_bytes()).hexdigest()==before
            store=json.loads((root/'daily_data/weeks/2026-11-02.json').read_text())
            assert '2026-11-04' not in store['days']
        update(4)
        # New Monday also catches up the previous Thursday/Friday into their old workbook.
        update(9)
        second=d.workbook_path(root,dt.date(2026,11,9))
        assert second.exists() and first.exists()
        for file in [first,second]:
            book=load_workbook(file,data_only=True)
            for name in book.sheetnames:
                if name in {'Sources','CPR Guide'}: continue
                day=dt.datetime.strptime(name,'%d-%b-%Y').date()
                expected=d.collect_day(day,True)['rows']
                got=[list(r) for r in book[name].iter_rows(min_row=2,max_row=len(expected)+1,max_col=5,values_only=True)]
                assert got==expected,(file,name)
        assert load_workbook(first).sheetnames==['02-Nov-2026','03-Nov-2026','04-Nov-2026','05-Nov-2026','06-Nov-2026','Sources','CPR Guide']
print('PASS: real exports; day addition; no duplicate rewrite; locked-file preservation and retry; missed-day catch-up; week rollover; every source price reconciled.')
