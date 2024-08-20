# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 10:29:52 2024

@author: CWilson
"""


import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')
from insertGroupHoursToSQL import insertGroupHours
import datetime

print('Running insertGroupHoursToSQL_REMEDIATION...')

now = datetime.datetime.now()
two_day = (now - datetime.timedelta(days=2)).strftime('%m/%d/%Y')
four_day = (now - datetime.timedelta(days=4)).strftime('%m/%d/%Y')
ten_day =  (now - datetime.timedelta(days=10)).strftime('%m/%d/%Y')


try:
    x = insertGroupHours(two_day)
    x.doStuff()
except Exception as e:
    print(f'Could not complete insertGroupHours("{two_day}") \n {e}')
    
    
try:
    x = insertGroupHours(four_day)
    x.doStuff()
except Exception as e:
    print(f'Could not complete insertGroupHours("{four_day}") \n {e}')
    

try:
    x = insertGroupHours(ten_day)
    x.doStuff()
except Exception as e:
    print(f'Could not complete insertGroupHours("{ten_day}") \n {e}')