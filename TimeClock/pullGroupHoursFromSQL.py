# -*- coding: utf-8 -*-
"""
Created on Wed Aug 14 15:34:12 2024

@author: CWilson
"""


import pandas as pd
from sqlalchemy import create_engine, MetaData, Table, select, func, and_, DateTime
from sqlalchemy.orm import sessionmaker
from initSQLConnectionEngine import yield_SQL_engine
import datetime


'''
def return_sql_times_df(date_str):
    engine = yield_SQL_engine()
    metadata = MetaData()
    # log_table = Table('employeeinformation_log', metadata, autoload_with=engine, schema='dbo')
    
    
    Session = sessionmaker(bind=engine)
    session = Session()
    
    # # first up, we need to check for when the last completed merge proc was performed
    # #select max(insertedat) from dbo.employeeinfomration_log
    # # stmt = select()
    # mostRecentMerge = session.query(func.max(log_table.c.insertedat)).scalar()
    
    # with engine.connect() as connection:
    #     result = connection.execute(text(f"select name, jobcode, costcode, hours, timein, timeout, isclockedin from dbo.clocktimes where timein::timestamp::date={date_str}"))
    #     for row in result:
    #         continue
    
    # mostRecentMerge = row[0]
    
    
    # # use this to check if the last merge was within valid time frame
    # if (datetime.datetime.now() - mostRecentMerge).seconds > 24*60*60*2: #greater than 2 days old:
    #     raise Exception('EmployeeInformationOlderThan2days')
    
    
    clocktimes_table = Table('clocktimes', metadata, autoload_with=engine, schema='dbo')
    query = session.query(clocktimes_table)
    result = query.all()
    
    # Convert the query result to a list of dictionaries
    data = [row._asdict() for row in result]
    
    # Create a DataFrame from the list of dictionaries
    df = pd.DataFrame(data)
    
    return df
'''



def return_sql_times_df(date_str):
    engine = yield_SQL_engine()
    
    # Reflect the table from the database
    metadata = MetaData()
    clocktimes_table = Table('clocktimes', metadata, autoload_with=engine, schema='dbo')
    
    # Create a session
    Session = sessionmaker(bind=engine)
    session = Session()
    
    # Define your date string
    date_str = '2024-08-13'
    
    # Create the subquery
    subquery = (
        session.query(
            func.cast(func.date_trunc('day', clocktimes_table.c.timein), DateTime).label('indate'),
            func.max(clocktimes_table.c.remediationtype).label('maxremediation')
        )
        .group_by(func.cast(func.date_trunc('day', clocktimes_table.c.timein), DateTime))
        .subquery()
    )
    
    # Main query
    query = (
        session.query(clocktimes_table)
        .join(
            subquery,
            and_(
                subquery.c.indate == func.cast(func.date_trunc('day', clocktimes_table.c.timein), DateTime),
                subquery.c.maxremediation == clocktimes_table.c.remediationtype
            )
        )
        .filter(func.cast(func.date_trunc('day', clocktimes_table.c.timein), DateTime) == date_str)
        .order_by(clocktimes_table.c.name, clocktimes_table.c.timein)
    )
    
    # Convert the query to a SQL string
    sql_query = query.statement
    
    # Load the results into a Pandas DataFrame
    df = pd.read_sql(sql_query, con=engine)
    
    # Display the DataFrame
    print(df)    
    