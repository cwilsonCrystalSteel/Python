# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:17:22 2024

@author: CWilson
"""

import pandas as pd
from sqlalchemy import MetaData, Table, func
from sqlalchemy.orm import sessionmaker
from initSQLConnectionEngine import yield_SQL_engine
import datetime


engine = yield_SQL_engine()
metadata = MetaData()
log_table = Table('employeeinformation_log', metadata, autoload_with=engine, schema='dbo')


Session = sessionmaker(bind=engine)
session = Session()

# first up, we need to check for when the last completed merge proc was performed
#select max(insertedat) from dbo.employeeinfomration_log
# stmt = select()
mostRecentMerge = session.query(func.max(log_table.c.insertedat)).scalar()

# use this to check if the last merge was within valid time frame
(datetime.datetime.now() - mostRecentMerge).seconds


employee_table = Table('employeeinformation', metadata, autoload_with=engine, schema='dbo')
query = session.query(employee_table)
result = query.all()

# Convert the query result to a list of dictionaries
data = [row._asdict() for row in result]

# Create a DataFrame from the list of dictionaries
df = pd.DataFrame(data)

