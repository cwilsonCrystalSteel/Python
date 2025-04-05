# -*- coding: utf-8 -*-
"""
Created on Fri Mar 28 17:20:48 2025

@author: Netadmin
"""

from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
from utils.sql_print_count_results import table_exists, print_count_results
import pandas as pd
from Google.pullEVALotsLogFromGoogle import pullEVALotsLogFromGoogle


    


def import_Lots_Log_EVA_hours_to_SQL(source=None):
    table = 'evalotslog'
    print('Retrieving LOTS Log from google sheets...')
    ll = pullEVALotsLogFromGoogle(drop_nonnumeric=False)
    
    engine = yield_SQL_engine()
    
    
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before truncating')
    
    if table_exists(engine, schema='live', table=table):
        Session = sessionmaker(bind=engine)
        session = Session()
        session.execute(text(f'''TRUNCATE TABLE live.{table}'''))
        session.commit()
        session.close()
    
    
    # get count of table before insert --> should be 0
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before importing')
    print_count_results(engine=engine, schema='dbo', table=table, suffix_text='before merge proc')
    
    ll.to_sql(table, engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.merge_evalotslog()")
    else:
        cursor.execute("call dbo.merge_evalotslog(%s)", (source,))
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine=engine, schema='dbo', table=table, suffix_text='after merge proc') 
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after merge proc')
