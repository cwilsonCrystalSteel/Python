# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 12:29:10 2024

@author: CWilson
"""


import pandas as pd
from utils.initSQLConnectionEngine import yield_SQL_engine
from sqlalchemy import create_engine, MetaData, Table, select, func, and_, DateTime
from sqlalchemy.orm import sessionmaker

def return_sql_jobcostcodes():
    engine = yield_SQL_engine()
    
    
    engine = yield_SQL_engine()
    metadata = MetaData()
    jcc_table = Table('job_costcodes', metadata, autoload_with=engine, schema='dbo')
    Session = sessionmaker(bind=engine)
    session = Session()   
    
    query = (
        select(jcc_table)  # No square brackets around clocktimes
    )
    # Execute the query
    result = session.execute(query)
    # Convert to DataFrame
    df = pd.DataFrame(result.fetchall(), columns=result.keys())    
    
    
    return df

