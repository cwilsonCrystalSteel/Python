# -*- coding: utf-8 -*-
"""
Created on Sun Apr  6 10:41:48 2025

@author: Netadmin
"""


import pandas as pd
from sqlalchemy import create_engine, MetaData, Table, select, func, and_, DateTime
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
import datetime
from pathlib import Path

def get_fablisting_from_vreturnevafromfablisting():
            
    engine = yield_SQL_engine()
    metadata = MetaData()
    vreturnevafromfablisting = Table('vreturnevafromfablisting', metadata, autoload_with=engine, schema='dbo')
    
    # Create a session
    Session = sessionmaker(bind=engine)
    
    with Session() as session:
       
        query = (
            select(vreturnevafromfablisting)
        )    
        
        
        result = session.execute(query)
        
        # Convert to DataFrame
        fablisting_eva = pd.DataFrame(result.fetchall(), columns=result.keys())    
        
    return fablisting_eva