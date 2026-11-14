"""Shared, locale-independent report names and calendar-month output folders."""
import datetime as dt
from pathlib import Path

MONTHS=('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')

def month_folder(output_root,period_start):
    return Path(output_root)/f'{period_start.year:04d}-{period_start.month:02d}'

def weekly_filename(start):
    end=start+dt.timedelta(days=4);a=MONTHS[start.month-1];b=MONTHS[end.month-1]
    if start.year!=end.year:label=f'{a} {start.day} {start.year}-{b} {end.day} {end.year}'
    elif start.month!=end.month:label=f'{a} {start.day}-{b} {end.day} {start.year}'
    else:label=f'{a} {start.day}-{end.day} {start.year}'
    return 'Weekly '+label+'.xlsx'

def weekly_path(root,start,output_root=None):
    return month_folder(output_root or Path(root)/'outputs',start)/weekly_filename(start)

def monthly_path(root,start):
    return month_folder(Path(root)/'outputs',start)/f'Monthly {MONTHS[start.month-1]} {start.year}.xlsx'

def daily_path(root,monday):
    end=monday+dt.timedelta(days=4)
    return month_folder(Path(root)/'outputs',monday)/f'NSE_FO_Daily_OHLC_{monday}_to_{end}.xlsx'
