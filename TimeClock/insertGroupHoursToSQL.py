# -*- coding: utf-8 -*-
"""
Created on Thu Aug  8 15:40:04 2024

@author: CWilson
"""


from sqlalchemy import text
import pandas as pd
from TimeClockNavigation import TimeClockBase, TimeClockEZGroupHours
from initSQLConnectionEngine import yield_SQL_engine
from Read_Group_hours_HTML import new_and_imporved_group_hours_html_reader
import os
import datetime

engine = yield_SQL_engine()


def print_count_results(table_name, engine, suffix_text):
    with engine.connect() as connection:
        result = connection.execute(text(f"select count(*) from live.{table_name}"))
        for row in result:
            continue
    
    print(f"There are {row[0]} rows in live.{table_name} {suffix_text}")
    

def getDatesTimesDF(date_str):
    i = 0
    while i < 5:
        try:
            x = TimeClockEZGroupHours(date_str)
            filepath = x.get_filepath()
            x.kill()
            times_df = new_and_imporved_group_hours_html_reader(filepath, in_and_out_times=True)
            
            os.remove(filepath)
            break
        except:
            print('oh no we failed on this attemp')
            i += 1
            try: 
                x.kill()
            except:
                print('no x to kill')
    
    return times_df
    

def insertInstananeous(times_df):
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call live.mergeclocktimes_fromstaging()")
    connection.commit()
    connection.close()
    
    return None

def insertRemediated(remediation_days=1):
    
    now = datetime.datetime.now()
    date_str = (now - datetime.timedelta(days=remediation_days)).strftime('%m/%d/%Y')
    times_df = getDatesTimesDF(date_str)
    
    table_name = 'clocktimes_remediation'
    
    print_count_results(table_name, engine, 'before truncating')
    connection = engine.connect()
    connection.execute(f"TRUNCATE TABLE live.{table_name}")
    connection.close()
    print_count_results(table_name, engine, 'after truncating')
    
    times_df = times_df.rename(columns={'Name':'name', 
                                        'Job Code':'jobcode',
                                        'Cost Code': 'costcode',
                                        'Hours': 'hours',
                                        'Time In': 'timein', 
                                        'Time Out': 'timeout'})
    times_df['timeout'] = pd.to_datetime(times_df['timeout'], errors='coerce')
    
    
    # NEED TO FIGURE OUT WHAT TO DO IF WE HAVE REMEDIATED HORUS THAT ARE CLOCKED IN #???
    times_df = times_df[~times_df['timeout'].isna()]
    
    times_df = times_df[~times_df['timein'].isna()]
    
    times_df.loc[:,'remediationtype'] = remediation_days
    
    # need to have a check for when times_df is empty --- it might not make it this far if that happens
    
    times_df.to_sql(table_name, engine, schema='live', if_exists='append', index=False)
    
    print_count_results(table_name, engine, 'after inserting')
        
    
    
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call dbo.mergeclocktimes()")
    connection.commit()
    connection.close()
    
    
    print_count_results(table_name, engine, 'after merge proc') # should be 0
    
    return None


def get_a_bunch_thisisaoneoff():
    daysback = 2
    daysbacktoo = 100
    for i in range(daysback, daysbacktoo):
        insertRemediated(remediation_days=i)