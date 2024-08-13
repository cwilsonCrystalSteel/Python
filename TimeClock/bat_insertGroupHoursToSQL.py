# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 13:54:25 2024

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')
from insertGroupHoursToSQL import insertRemediated

print('Running import_employee_information_to_SQL...')

try:
    insertRemediated(1)
except Exception as e:
    print(f'Could not complete insertRemediated(1) \n {e}')

try:
    insertRemediated(3)
except Exception as e:
    print(f'Could not complete insertRemediated(3) \n {e}')
    
try:
    insertRemediated(10)
except Exception as e:
    print(f'Could not complete insertRemediated(5) \n {e}')