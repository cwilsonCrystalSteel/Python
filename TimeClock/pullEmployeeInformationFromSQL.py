# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:17:22 2024

@author: CWilson
"""

import pandas as pd
from sqlalchemy import MetaData, Table, func, text
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
import datetime


def return_sql_ei():
    engine = yield_SQL_engine()
    metadata = MetaData()
    
    
    Session = sessionmaker(bind=engine)
    session = Session()
    
    # # first up, we need to check for when the last completed merge proc was performed
    # #select max(insertedat) from dbo.employeeinfomration_log
    # # stmt = select()
    # mostRecentMerge = session.query(func.max(log_table.c.insertedat)).scalar()
    
    with engine.connect() as connection:
        result = connection.execute(text("select max(insertedat) from dbo.employeeinformation_log where description='merge_employeeinformation'"))
        for row in result:
            continue
    
    mostRecentMerge = row[0]
    
    
    # use this to check if the last merge was within valid time frame
    if (datetime.datetime.now() - mostRecentMerge).seconds > 24*60*60*2: #greater than 2 days old:
        raise Exception('EmployeeInformationOlderThan2days')
    
    
    employee_table = Table('employeeinformation', metadata, autoload_with=engine, schema='dbo')
    query = session.query(employee_table)
    result = query.all()
    
    # Convert the query result to a list of dictionaries
    data = [row._asdict() for row in result]
    
    # Create a DataFrame from the list of dictionaries
    df = pd.DataFrame(data)
    
    return df

