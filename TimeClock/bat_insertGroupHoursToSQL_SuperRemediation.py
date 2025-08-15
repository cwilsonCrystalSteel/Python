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

source = 'bat_insertGroupHoursToSQL_SuperRemediation'

for i in range(2,60):
    
    date = (now - datetime.timedelta(days=i)).strftime('%m/%d/%Y')
    
    
        
    try:
        download_folder = Path.home() / 'downloads' / 'GroupHours_SuperRemediation'
        x = insertGroupHours(date_str=date, download_folder=download_folder, source=source, headless=True, offscreen=False)
        x.doStuff()
    except Exception as e:
        output_error_string = f'Could not complete insertGroupHours("{date}") \n {e}'
        print(output_error_string)
        insertError(name=source, description = output_error_string)
    
    