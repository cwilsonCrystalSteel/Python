# -*- coding: utf-8 -*-
"""
Created on Sat Apr  5 12:57:06 2025

@author: Netadmin
"""


from Google.insertHPTProductionSheetToSQL import import_ProudctionSheet_HPT_hours_to_SQL
from utils.insertErrorToSQL import insertError

source = 'bat_insertHPTProductionSheetToSQL'


try:
    import_ProudctionSheet_HPT_hours_to_SQL(source=source)
except Exception as e:
    insertError(name=source, description = str(e))
    print(f'Could not complete import_ProudctionSheet_HPT_hours_to_SQL\n {e}')
