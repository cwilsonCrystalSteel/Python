# -*- coding: utf-8 -*-
"""
Created on Sun Mar 30 12:49:54 2025

@author: Netadmin
"""


import pandas as pd
from sqlalchemy import create_engine, MetaData, Table, select, func, and_, DateTime
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
import datetime


def get_sql_evaDropbox():
    engine = yield_SQL_engine()
    metadata = MetaData()
    dropbox_eva_table = Table('evadropbox', metadata, autoload_with=engine, schema='dbo')
    return dropbox_eva_table

def return_select_evadropbox_where_filename(excel_file_str=''):
    engine = yield_SQL_engine()
    dropbox_eva_table = get_sql_evaDropbox()
    
    
    Session = sessionmaker(bind=engine)
    with Session() as session:
    
        # Create the main query
        query = (
            select(dropbox_eva_table)
            .where(dropbox_eva_table.c.filepath == excel_file_str)
        )    
    
        result = session.execute(query)
        
        # Convert to DataFrame
        eva_df = pd.DataFrame(result.fetchall(), columns=result.keys())    

    return eva_df


def return_select_evadropbox_all():
    engine = yield_SQL_engine()
    dropbox_eva_table = get_sql_evaDropbox()
    
    
    Session = sessionmaker(bind=engine)
    with Session() as session:
        query = (
            select(dropbox_eva_table)
        )

        result = session.execute(query)
        
        # Convert to DataFrame
        eva_df = pd.DataFrame(result.fetchall(), columns=result.keys())    

    return eva_df        