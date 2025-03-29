# -*- coding: utf-8 -*-
"""
Created on Sat Mar 29 11:51:58 2025

@author: Netadmin
"""

from Google.insertEVALotsLogToSQL import import_Lots_Log_EVA_hours_to_SQL
from utils.insertErrorToSQL import insertError

source = 'bat_insertEVALotsLogToSQL'


try:
    import_Lots_Log_EVA_hours_to_SQL(source=source)
except Exception as e:
    insertError(name=source, description = str(e))
    print(f'Could not complete import_Lots_Log_EVA_hours_to_SQL\n {e}')
