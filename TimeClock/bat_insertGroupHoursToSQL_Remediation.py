# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 10:29:52 2024

@author: CWilson
"""


from TimeClock.insertGroupHoursToSQL import insertGroupHours
import datetime
from pathlib import Path
from utils.insertErrorToSQL import insertError

print('Running insertGroupHoursToSQL_REMEDIATION...')

now = datetime.datetime.now()
two_day = (now - datetime.timedelta(days=2)).strftime('%m/%d/%Y')
four_day = (now - datetime.timedelta(days=4)).strftime('%m/%d/%Y')
ten_day =  (now - datetime.timedelta(days=10)).strftime('%m/%d/%Y')

source = 'bat_insertGroupHoursToSQL_Remediation'



try:
    download_folder = Path.home() / 'downloads' / 'GroupHours_2DayRemediation'
    x = insertGroupHours(date_str=two_day, download_folder=download_folder, source=source, headless=True, offscreen=False)
    x.doStuff()
except Exception as e:
    output_error_string = f'Could not complete insertGroupHours("{two_day}") \n {e}'
    print(output_error_string)
    insertError(name=source, description = output_error_string)
    
    
try:
    download_folder = Path.home() / 'downloads' / 'GroupHours_4DayRemediation'
    x = insertGroupHours(date_str=four_day, download_folder=download_folder, source=source, headless=True, offscreen=False)
    x.doStuff()
except Exception as e:
    output_error_string = f'Could not complete insertGroupHours("{four_day}") \n {e}'
    print(output_error_string)
    insertError(name=source, description = output_error_string)
    

try:
    download_folder = Path.home() / 'downloads' / 'GroupHours_10DayRemediation'
    x = insertGroupHours(date_str=ten_day, download_folder=download_folder, source=source, headless=True, offscreen=False)
    x.doStuff()
except Exception as e:
    output_error_string = f'Could not complete insertGroupHours("{ten_day}") \n {e}'
    print(output_error_string)
    insertError(name=source, description = output_error_string)