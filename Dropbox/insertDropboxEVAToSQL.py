# -*- coding: utf-8 -*-
"""
Created on Sun Mar 30 12:45:26 2025

@author: Netadmin
"""


from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
import pandas as pd
from utils.initSQLConnectionEngine import yield_SQL_engine
from utils.sql_print_count_results import print_count_results, table_exists
from pathlib import Path

engine = yield_SQL_engine()

def insert_evaDropbox_log(description='', source=''):
    Session = sessionmaker(bind=engine)
    session = Session()

    query = text("""
    INSERT INTO dbo.eva_log (description, insertedat, source)
    VALUES (:description, NOW(), :source)
    """)
    session.execute(query, {"description": description, "source": source})
    session.commit()
    

    
def import_dropbox_eva_to_SQL(excel_lot_df, source=None):
    table = 'evadropbox'
    
    print_count_results(engine, schema='live', table=table, suffix_text='before truncating')
    
    if table_exists(engine, 'live', table):
        Session = sessionmaker(bind=engine)
        session = Session()
        session.execute(text(f'''TRUNCATE TABLE live.{table}'''))
        session.commit()
        session.close()
    
    
    # get count of table before insert --> should be 0
    print_count_results(engine, schema='live', table=table, suffix_text='before importing')
    print_count_results(engine, schema='dbo', table=table, suffix_text='before merge proc')
    
    excel_lot_df.to_sql(table, engine, schema='live', if_exists='replace', index=False)
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
    
    
def delete_filepath_for_redo(filepath_str):
    table = 'evadropbox'
    
    if not isinstance(filepath_str, str):
        filepath_str = str(filepath_str)
        
        
    delete_query = f"delete from dbo.{table} where filepath = '{filepath_str}'"
    
    
    print_count_results(engine, schema='dbo', table=table, suffix_text='before deleting filepath')
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute(delete_query)
    connection.commit()
    connection.close()
    
    print_count_results(engine, schema='dbo', table=table, suffix_text='after deleting filepath')
        
        