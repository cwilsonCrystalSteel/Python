# -*- coding: utf-8 -*-
"""
Created on Thu Feb 27 10:22:37 2025

@author: CWilson
"""

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

print('Running MANUAL_REMEDIATION_insertGroupHoursToSQL...')



now = datetime.datetime.now()

dates_to_try = ['2024-02-08', '2024-02-09', '2024-02-10', '2024-02-11', '2024-02-12', 
                '2024-02-12', '2024-02-13', '2024-02-01', '2024-02-03', '2024-02-06', 
                '2024-02-18', '2024-02-19', '2024-02-20']

for i in dates_to_try:
    
    date_str_eligible = datetime.datetime.strptime(i, '%Y-%m-%d').strftime('%m/%d/%Y')
    

    source = 'MANUAL_REMEDIATION_insertGroupHoursToSQL'

    
    try:
        x = insertGroupHours(date_str=date_str_eligible, download_folder=r'c:\users\cwilson\downloads\groupHours_manualRemediation')
        x.doStuff()
    except Exception as e:
        print(f'Could not complete insertGroupHours("{date_str_eligible}") \n {e}')
        