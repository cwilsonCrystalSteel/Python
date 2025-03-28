# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 08:59:57 2024

@author: CWilson
"""

from TimeClock.insertGroupHoursToSQL import insertGroupHours
import datetime
from pathlib import Path

print('Running insertGroupHoursToSQL_TODAY...')

now = datetime.datetime.now()
today_str = now.strftime('%m/%d/%Y')

source = 'bat_insertGroupHoursToSQL_Today'

download_folder = Path.home() / 'downloads' / 'GroupHours_Today'


try:
    x = insertGroupHours(date_str=today_str, download_folder=download_folder, source=source, headless=True, offscreen=False)
    x.doStuff()
except Exception as e:
    print(f'Could not complete insertGroupHours("{today_str}") \n {e}')
