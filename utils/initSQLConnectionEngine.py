# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:19:04 2024

@author: CWilson
"""

from sqlalchemy import create_engine
from utils.sqlCredentials import returnSqlCredentials


def yield_SQL_engine():
    
    sql_creds = returnSqlCredentials()
    username = sql_creds['username']
    password = sql_creds['password']
    
    
    engine = create_engine(f'postgresql://{username}:{password}@localhost:5432/postgres')
    return engine