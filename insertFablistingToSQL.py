# -*- coding: utf-8 -*-
"""
Created on Thu Apr  3 14:59:28 2025

@author: Netadmin
"""
''' get_model_estimate_hours_attached_to_fablisting_SQL '''
from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
import pandas as pd
from utils.initSQLConnectionEngine import yield_SQL_engine, table_exists
from utils.sql_print_count_results import print_count_results
from pathlib import Path

engine = yield_SQL_engine()


def further_fablisting_df_preperations(fablisting_df): # move this to get_model_estimate_hours_attached_to_fablisting_SQL
    df = fablisting_df.copy()
    
    df['pcmark'] = df['Piece Mark - REV'].apply(lambda x: x.split('-')[0])
    df['rev'] = df['Piece Mark - REV'].apply(lambda x: x.split('-')[1] if '-' in x else 0)
    
    df['lot_3_digit'] = df['Lot #'].str.zfill(3)
    df['lot_with_t_start'] = 'T' + df['lot_3_digit']
    
    cols = df.columns
    new_cols = []
    counter = 0
    for i in cols:
        if i == '' or i.isspace() or not i:
            new_cols.append(f'column {counter}')
            counter += 1
        else:
            new_cols.append(i)
    df.columns = new_cols
            
    df['Job #'] = df['Job #'].astype(str)
    
    return df


def import_dropbox_eva_to_SQL(fablisting, source=None):
    fablisting_df = further_fablisting_df_preperations(fablisting) # move this to get_model_estimate_hours_attached_to_fablisting_SQL
    table = 'fablisting'
    
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
    
    fablisting_df.to_sql(table, engine, schema='live', if_exists='replace', index=False)
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
