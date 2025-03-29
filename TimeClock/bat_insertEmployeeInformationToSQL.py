# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:01:47 2024

@author: CWilson
"""

from TimeClock.insertEmployeeInformationToSQL import import_employee_information_to_SQL, determine_terminated_employees
from utils.insertErrorToSQL import insertError

source = 'bat_insertEmployeeInformationToSQL'

print('Running import_employee_information_to_SQL...')

try:
    import_employee_information_to_SQL(source=source)
except Exception as e:
    insertError(name=source, description = str(e))
    print(f'Could not complete import_employee_information_to_SQL because of \n {e}')

try:
    determine_terminated_employees(source=source)
except Exception as e:
    insertError(name=source, description = str(e))
    print(f'Could not complete determine_terminated_employees because of \n {e}')
    