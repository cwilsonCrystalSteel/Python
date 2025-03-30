# -*- coding: utf-8 -*-
"""
Created on Sun Mar 30 12:45:26 2025

@author: Netadmin
"""


from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
import pandas as pd
from utils.initSQLConnectionEngine import yield_SQL_engine
from utils.sql_print_count_results import print_count_results
from pathlib import Path

engine = yield_SQL_engine()



    
def import_dropbox_eva_to_SQL(excel_lot_df, source=None):
    table = 'evadropbox'
    
    print_count_results(engine, schema='live', table=table, suffix_text='before truncating')
    
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text(f'''TRUNCATE TABLE live.{table}'''))
    session.commit()
    session.close()
    
    
    # get count of table before insert --> should be 0
    print_count_results(engine, schema='live', table=table, suffix_text='before importing')
    print_count_results(engine, schema='dbo', table=table, suffix_text='before merge proc')
    
    excel_lot_df.to_sql('evadropbox', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine, schema='live', table=table, suffix_text='after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.merge_evadropbox()")
    else:
        cursor.execute("call dbo.merge_evadropbox(%s)", (source,))
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine, schema='dbo', table=table, suffix_text='after merge proc')
    print_count_results(engine, schema='live', table=table, suffix_text='after merge proc')