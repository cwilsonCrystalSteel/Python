# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:01:47 2024

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')
from insertEmployeeInformationToSQL import import_employee_information_to_SQL, determine_terminated_employees


print('Running import_employee_information_to_SQL...')

try:
    import_employee_information_to_SQL()
except Exception as e:
    print(f'Could not complete import_employee_information_to_SQL because of \n {e}')

try:
    determine_terminated_employees()
except Exception as e:
    print(f'Could not complete determine_terminated_employees because of \n {e}')