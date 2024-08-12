# -*- coding: utf-8 -*-
"""
Created on Thu Aug  8 15:40:04 2024

@author: CWilson
"""


from sqlalchemy import text
import pandas as pd
from TimeClockNavigation import TimeClockBase, TimeClockEZGroupHours
from initSQLConnectionEngine import yield_SQL_engine

engine = yield_SQL_engine()


def insertInstananeous(times_df):
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call live.mergeclocktimes_fromstaging()")
    connection.commit()
    connection.close()
    
    return None

def insertRemediated(times_df, remediation_days=1):
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call dbo.mergeclocktimes()")
    connection.commit()
    connection.close()
    return None