# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 12:29:10 2024

@author: CWilson
"""


import pandas as pd
from initSQLConnectionEngine import yield_SQL_engine


def return_sql_jobcostcodes():
    engine = yield_SQL_engine()
    
    jccSQL = pd.read_sql_table(table_name = 'job_costcodes', schema='dbo', con=engine)

    
    return jccSQL

