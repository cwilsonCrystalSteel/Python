# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 10:29:52 2024

@author: CWilson
"""

from TimeClock.pullGroupHoursFromSQL import get_timesdf_from_vClocktimes
from TimeClock.insertGroupHoursToSQL import insertGroupHours
import datetime
from pathlib import Path
import pandas as pd

print('Running insertGroupHoursToSQL_Manual...')

now = datetime.datetime.now()

source = 'bat_insertGroupHoursToSQL_Manual'

times_df = get_timesdf_from_vClocktimes('2022-01-01','2028-04-01')

noEmployeeNumber = times_df[times_df['employeeidnumber'].isna()]

noEmployeeNumber = noEmployeeNumber[noEmployeeNumber['source'] == 'clocktimes']

targetDates = pd.unique(noEmployeeNumber['targetdate'])


for i in targetDates:
    date_str = i.strftime('%m/%d/%Y')
        
    try:
        download_folder = Path.home() / 'downloads' / 'GroupHours_manual'
        x = insertGroupHours(date_str=date_str, download_folder=download_folder)
        x.doStuff()
    except Exception as e:
        print(f'Could not complete insertGroupHours("{date_str}") \n {e}')
    
    