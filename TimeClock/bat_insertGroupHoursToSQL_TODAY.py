# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 08:59:57 2024

@author: CWilson
"""


import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')
from insertGroupHoursToSQL import insertGroupHours
import datetime

print('Running insertGroupHoursToSQL_TODAY...')

now = datetime.datetime.now()
today_str = now.strftime('%m/%d/%Y')

source = 'bat_insertGroupHoursToSQL_Today'



try:
    x = insertGroupHours(date_str=today_str, download_folder=r'c:\users\cwilson\downloads\GroupHours_Today')
    x.doStuff()
except Exception as e:
    print(f'Could not complete insertGroupHours("{today_str}") \n {e}')
