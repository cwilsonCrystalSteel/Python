# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 10:28:21 2024

@author: CWilson
"""


import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')
from insertGroupHoursToSQL import insertGroupHours
import datetime

print('Running insertGroupHoursToSQL_YESTERDAY...')

now = datetime.datetime.now()
yesterday_str = (now - datetime.timedelta(days=1)).strftime('%m/%d/%Y')


try:
    x = insertGroupHours(yesterday_str)
    x.doStuff()
except Exception as e:
    print(f'Could not complete insertGroupHours("{yesterday_str}") \n {e}')
