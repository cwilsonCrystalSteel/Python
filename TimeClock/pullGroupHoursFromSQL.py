# -*- coding: utf-8 -*-
"""
Created on Wed Aug 14 15:34:12 2024

@author: CWilson
"""


import pandas as pd
from sqlalchemy import create_engine, MetaData, Table, select, func, and_, DateTime
from sqlalchemy.orm import sessionmaker
from utils.initSQLConnectionEngine import yield_SQL_engine
import datetime
from pathlib import Path

from TimeClock.insertGroupHoursToSQL import insertGroupHours


    
def get_today_or_yesterdays_timesdf(today_or_yesterday):
    if today_or_yesterday == 'today':
        tablename = 'clocktimes_today'
    elif today_or_yesterday == 'yesterday':
        tablename = 'clocktimes_yesterday'
        
    engine = yield_SQL_engine()
    metadata = MetaData()
    today_or_yesterday_table = Table(tablename, metadata, autoload_with=engine, schema='dbo')
    Session = sessionmaker(bind=engine)
    with Session() as session:
        
        query = (
            select(today_or_yesterday_table)  # No square brackets around clocktimes
        )
        # Execute the query
        result = session.execute(query)
        # Convert to DataFrame
        df = pd.DataFrame(result.fetchall(), columns=result.keys())
    
    
    return df



def date_str_handler(date_str):
    
    try:
        d = datetime.datetime.strptime(date_str, '%m/%d/%Y')
        # print('Date passed with format %m/%d/%Y')
        return d.strftime('%Y-%m-%d')
    except ValueError:
        pass
        # print('Date was not format %m/%d/%Y')
        
        
    try:
        d = datetime.datetime.strptime(date_str, '%Y-%m-%d')
        return date_str
    except Exception as e:
        print(f'Could not determine date_str format with error: {e}')

def check_remediated_availability(date_str):
    """
    This will query dbo.clocktimes to ensure we have atleast some records
    for the targetdate = date_str
    
    Parameters
    ----------
    date_str : str %Y-%m-%d or %m/%d/%Y
        select * from dbo.clocktimes where targetdate = date_str

    Returns
    -------
    bool
        can we proceed with the provided date? bool.

    """
    date_str_for_sql = date_str_handler(date_str)
    
    engine = yield_SQL_engine()
    metadata = MetaData()
    clocktimes = Table('clocktimes', metadata, autoload_with=engine, schema='dbo')
    Session = sessionmaker(bind=engine)
    with Session() as session: 
        
        query = (
            select(clocktimes)  # No square brackets around clocktimes
            .where(clocktimes.c.targetdate == date_str_for_sql)
        )
        # Execute the query
        result = session.execute(query)
        # Convert to DataFrame
        df = pd.DataFrame(result.fetchall(), columns=result.keys())    
        
    return bool(df.shape[0])



def format_remediated_like_YesterdayOrToday_tables(times_df):
    """
    This is an easy way to make sure the columns from dbo.clocktimes
    match results returned from dbo.clocktimes_today and dbo.clocktimes_yesterday

    Parameters
    ----------
    times_df : Pandas DataFrame
        an output of the sql table dbo.clocktimes.

    Returns
    -------
    times_df : Pandas DataFrame
        returns only select columns.

    """
    
    
    times_df = times_df[['employeeidnumber', 'job_costcode_id', 'hours', 'timein', 'timeout','isclockedin']]
    
    return times_df
    
 
def get_date_range_timesdf_controller(start_date, end_date):
    # determine the date range if it passes into yesterday and today
    # determine if date range exceeds into the future
    start_date_sql = date_str_handler(start_date)
    end_date_sql = date_str_handler(end_date)

    start_dt = datetime.datetime.strptime(start_date_sql, '%Y-%m-%d')
    end_dt = datetime.datetime.strptime(end_date_sql, '%Y-%m-%d')

    if start_date_sql == end_date_sql:
        # move the end date up one day if we get the same day for start & end date
        end_dt = end_dt + datetime.timedelta(days=1)
        end_date_sql = datetime.datetime.strftime(end_dt, '%Y-%m-%d')
        
        
    
    number_days = (end_dt - start_dt).days
    
    table_decider = {'clocktimes':[],
                     'clocktimes_yesterday':None,
                     'clocktimes_today':None,
                     'future':[]}
    
    today_dt = datetime.datetime.now().date()
    yesterday_dt = today_dt - datetime.timedelta(days=1)
    
    for i in range(0,number_days):
        checker_dt = (start_dt + datetime.timedelta(days=i)).date()
        if checker_dt == today_dt:
            table_decider['clocktimes_today'] = checker_dt
        elif checker_dt == yesterday_dt:
            table_decider['clocktimes_yesterday'] = checker_dt
        elif checker_dt > today_dt:
            table_decider['future'].append(checker_dt)
        else:
            table_decider['clocktimes'].append(checker_dt)
        
    
    
    to_union = []
    
    if len(table_decider['clocktimes']):
        range_start_date_str = min(table_decider['clocktimes']).strftime('%Y-%m-%d')
        range_end_date_str = max(table_decider['clocktimes']).strftime('%Y-%m-%d')
        
        range_df = get_date_range_timesdf_REMEDIATEDONLY(range_start_date_str, range_end_date_str)
        
        range_df = format_remediated_like_YesterdayOrToday_tables(range_df)
        
        to_union.append(range_df)
        
    if table_decider['clocktimes_yesterday'] is not None:
        yesterday_df = get_today_or_yesterdays_timesdf('yesterday')
        
        to_union.append(yesterday_df)
    
    if table_decider['clocktimes_today'] is not None:
        today_df = get_today_or_yesterdays_timesdf('today') 
        
        to_union.append(today_df)
        
    if len(table_decider['future']):
        future_min = min(table_decider['future']).strftime('%Y-%m-%d')
        future_max = max(table_decider['future']).strftime('%Y-%m-%d')
        print(f"Could not retrieve dates {future_min} to {future_max} because they're in the future!")


    output_df = pd.concat(to_union, axis=0)
    
    output_df['hours'] = pd.to_numeric(output_df['hours'], errors='coerce')
            
    return output_df

        
    
def get_date_range_timesdf_REMEDIATEDONLY(start_date, end_date):
    
    start_date_sql = date_str_handler(start_date)
    end_date_sql = date_str_handler(end_date)
    
    
    engine = yield_SQL_engine()
    metadata = MetaData()
    clocktimes = Table('clocktimes', metadata, autoload_with=engine, schema='dbo')
    
    # Create a session
    Session = sessionmaker(bind=engine)
    with Session() as session:
        
        subquery = (
            select(
                clocktimes.c.targetdate,
                func.max(clocktimes.c.remediationtype).label('remediationtype')
            )
            .group_by(clocktimes.c.targetdate)
        ).alias('z')
    
        # Create the main query
        query = (
            select(clocktimes)
            .select_from(clocktimes.join(subquery, 
                (subquery.c.remediationtype == clocktimes.c.remediationtype) &
                (clocktimes.c.targetdate == subquery.c.targetdate)
            ))
            .where(clocktimes.c.targetdate >= start_date_sql, clocktimes.c.targetdate <= end_date_sql)
        )    
        
        
        result = session.execute(query)
        
        # Convert to DataFrame
        times_df = pd.DataFrame(result.fetchall(), columns=result.keys())    

    return times_df
    
    

def get_specific_dates_timesdf(date_str):
    
    date_str_for_sql = date_str_handler(date_str)
    
    if not check_remediated_availability(date_str):
        print(f'Trying to pull TimeClock for: {date_str} now...')
        try:
            download_folder = Path.home() / 'downloads' / 'GroupHours'
            x = insertGroupHours(date_str=date_str, download_folder=download_folder, source='pullGroupHoursFromSQL')
            x.doStuff()
            if check_remediated_availability(date_str):
                print(f'Good news, we alleviated missing data on {date_str}')
            else:
                print(f'Bad news, we could not infill data for {date_str}')
        except Exception as e:
            print(f'Could not complete insertGroupHours("{date_str}") \n {e}')
            
        
        
    
    
    engine = yield_SQL_engine()
    metadata = MetaData()
    clocktimes = Table('clocktimes', metadata, autoload_with=engine, schema='dbo')
    
    # Create a session
    Session = sessionmaker(bind=engine)
    with Session() as session: 
        
        subquery = (
            select(func.max(clocktimes.c.remediationtype).label('remediationtype'))
            .where(clocktimes.c.targetdate == date_str_for_sql)
            .group_by(clocktimes.c.targetdate)
        ).alias('z')
        
     
        # Create the main query
        query = (
            select(clocktimes)  # No square brackets around clocktimes
            .select_from(clocktimes.join(subquery, subquery.c.remediationtype == clocktimes.c.remediationtype))
            .where(clocktimes.c.targetdate == date_str_for_sql)
        )
            
            
        # Execute the query
        result = session.execute(query)
        
        # Convert to DataFrame
        times_df = pd.DataFrame(result.fetchall(), columns=result.keys())    

    return times_df
    
    
def get_timesdf_from_sql(date_str):
    
    now = datetime.datetime.now().date()
    
    date_dt = datetime.datetime.strptime(date_str, '%m/%d/%Y').date()
    
    if date_dt == now:
        print("Querying the data from today's table")
        times_df = get_today_or_yesterdays_timesdf('today')
        
    elif (now - date_dt).days == 1:
        print("Querying the data from yesterday's table")
        times_df = get_today_or_yesterdays_timesdf('yesterday')
        
    else:
        print(f"Querying the data from the remediation table for date: {date_str}")
        times_df = get_specific_dates_timesdf(date_str)
    
    return times_df


def get_timesdf_from_vClocktimes(start_date, end_date):
    
    start_date_sql = date_str_handler(start_date)
    end_date_sql = date_str_handler(end_date)
        
    engine = yield_SQL_engine()
    metadata = MetaData()
    vclocktimes = Table('vclocktimes', metadata, autoload_with=engine, schema='dbo')
    
    # Create a session
    Session = sessionmaker(bind=engine)
    
    with Session() as session:
       
        query = (
            select(vclocktimes)
            .where(vclocktimes.c.targetdate >= start_date_sql, vclocktimes.c.targetdate <= end_date_sql)
        )    
        
        
        result = session.execute(query)
        
        # Convert to DataFrame
        times_df = pd.DataFrame(result.fetchall(), columns=result.keys())    
        
    return times_df


    
    