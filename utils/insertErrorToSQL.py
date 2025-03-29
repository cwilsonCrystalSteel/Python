# -*- coding: utf-8 -*-
"""
Created on Tue Sep 24 09:55:43 2024

@author: CWilson
"""

from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine


def insertError(name="", description=""):
    engine = yield_SQL_engine()
    
    # replace single quote with 2x single quote
    name = name.replace("'", "''")
    description = description.replace("'","''")
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text(f"INSERT INTO dbo.error_log (name, description) values ('{name}', '{description}')"))
    session.commit()
    session.close()
    
    return None