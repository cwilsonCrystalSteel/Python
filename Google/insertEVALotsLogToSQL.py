# -*- coding: utf-8 -*-
"""
Created on Fri Mar 28 17:20:48 2025

@author: Netadmin
"""

from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
import pandas as pd
from Google.pullEVALotsLogFromGoogle import pullEVALotsLogFromGoogle



def print_count_results(schema, engine, suffix_text):
    with engine.connect() as connection:
        result = connection.execute(text(f"select count(*) from {schema}.evaLotsLog"))
        for row in result:
            continue
    
    print(f"There are {row[0]} rows in {schema}.evaLotsLog {suffix_text}")
    


def import_Lots_Log_EVA_hours_to_SQL(source=None):
    print('Retrieving LOTS Log from google sheets...')
    ll = pullEVALotsLogFromGoogle(drop_nonnumeric=False)
    
    
    engine = yield_SQL_engine()
    
    
    print_count_results('live', engine, 'before truncating')
    
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.evalotslog'''))
    session.commit()
    session.close()
    
    
    # get count of table before insert --> should be 0
    print_count_results('live', engine, 'before importing')
    print_count_results('dbo', engine, 'before merge proc')
    
    ll.to_sql('evalotslog', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results('live', engine, 'after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.merge_evalotslog()")
    else:
        cursor.execute("call dbo.merge_evalotslog(%s)", (source,))
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results('dbo', engine, 'after merge proc') 
    print_count_results('live', engine, 'after merge proc')
