# -*- coding: utf-8 -*-
"""
Created on Sat Apr  5 12:51:58 2025

@author: Netadmin
"""


from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
from utils.sql_print_count_results import table_exists, print_count_results
import pandas as pd
from Google.pullHPTProductionSheetFromGoogle import pullProductionSheetFromGoogle


    


def import_ProudctionSheet_HPT_hours_to_SQL(source=None):
    table = 'hptproductionsheet'
    print('Retrieving Production Sheet from google sheets...')
    ps = pullProductionSheetFromGoogle()
    
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
    # print_count_results(engine=engine, schema='dbo', table=table, suffix_text='before merge proc')
    
    ps.to_sql(table, engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.merge_hptproductionsheet()")
    else:
        cursor.execute("call dbo.merge_hptproductionsheet(%s)", (source,))
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine=engine, schema='dbo', table=table, suffix_text='after merge proc') 
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after merge proc')
    # source = source if source is not None else ''
    # source = source.replace("'", "''")
    # description = 'insert into live.hptproductionsheet'
    
    # Session = sessionmaker(bind=engine)
    # session = Session()
    # session.execute(text(f"INSERT INTO dbo.eva_log (source, description, insertedat) values ('{source}', '{description}', now())"))
    # session.commit()
    # session.close()
