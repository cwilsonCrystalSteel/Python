# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 10:28:21 2024

@author: CWilson
"""

from TimeClock.insertGroupHoursToSQL import insertGroupHours
import datetime
from pathlib import Path
from utils.insertErrorToSQL import insertError

print('Running insertGroupHoursToSQL_YESTERDAY...')

now = datetime.datetime.now()
yesterday_str = (now - datetime.timedelta(days=1)).strftime('%m/%d/%Y')

source = 'bat_insertGroupHoursToSQL_Yesterday'



try:
    download_folder = Path.home() / 'downloads' / 'GroupHours_Yesterday'
    x = insertGroupHours(date_str=yesterday_str, download_folder=download_folder, source=source, headless=True, offscreen=False)
    x.doStuff()
except Exception as e:
    output_error_string = f'Could not complete insertGroupHours("{yesterday_str}") \n {e}'
    print(output_error_string)
    insertError(name=source, description = output_error_string)
