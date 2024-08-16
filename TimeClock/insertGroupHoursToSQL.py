# -*- coding: utf-8 -*-
"""
Created on Thu Aug  8 15:40:04 2024

@author: CWilson
"""


from sqlalchemy import create_engine, MetaData, Table, select, func, and_, DateTime, text
import pandas as pd
from TimeClockNavigation import TimeClockBase, TimeClockEZGroupHours
from initSQLConnectionEngine import yield_SQL_engine
from Read_Group_hours_HTML import new_and_imporved_group_hours_html_reader
import os
import datetime

engine = yield_SQL_engine()


def print_count_results(schema, table_name, engine, suffix_text):
    with engine.connect() as connection:
        result = connection.execute(text(f"select count(*) from {schema}.{table_name}"))
        for row in result:
            continue
    
    print(f"There are {row[0]} rows in {schema}.{table_name} {suffix_text}")



class CheckpointNotReachedException(Exception):
    def __init__(self, message=''):
        # Call the base class constructor with the parameters it needs
        super(CheckpointNotReachedException, self).__init__(message)



'''
x = insertGroupHours('08/13/2024')
x.getDatesTimesDF()
x.check_job_costcodes()
x.insertGroupHours()

x.doStuff()

'''

class insertGroupHours():
    def __init__(self, date_str, remediationtype=1):
        self.date_str = date_str
        self.remediationtype = remediationtype
        self.df_renamed_toggle = False
        
        self.engine = yield_SQL_engine()
        
        
        '''
        
        Before any merging occurs, we need to run check_job_costcodes
        
        '''
        self.job_costcodes_checked = False
        self.mergeTodayAvailable = False
        self.mergeYesterdayAvailable = False
        self.mergeRemediationAvailable = False
        
        today_dt = datetime.datetime.now().date()
        yesterday_dt = today_dt - datetime.timedelta(days=1)
        date_dt = datetime.datetime.strptime(self.date_str, '%m/%d/%Y').date()
        
        
        if date_dt == today_dt:
            self.mergeTodayAvailable = True
        elif date_dt == yesterday_dt:
            self.mergeYesterdayAvailable = True
        elif date_dt < yesterday_dt:
            self.mergeRemediationAvailable = True
            self.remediationtype = (today_dt - date_dt).days
            
        
    def doStuff(self):
        
        self.getDatesTimesDF()
        self.check_job_costcodes()
        self.insertGroupHours()
        
        
    def compare_sql_to_timeclock(self):
        
        '''
        I need this as a debugging tool
        
        Pass it a date_str, and it will return the best result of the sql with the timeclock
        
        it will also return all of the different remediation day results
        '''
        
        # get a fresh run of the date from timeclock
        self.getDatesTimesDF()
        self.timeclock = self.times_df.copy()
        
        # query and get the best result
        self.best_result = 5
        
        self.other_results = 10
        
        
        return None
        
        
      

    def getDatesTimesDF(self):
        i = 0
        while i < 5:
            try:
                x = TimeClockEZGroupHours(self.date_str)
                filepath = x.get_filepath()
                x.kill()
                if isinstance(filepath, Exception):
                    print(f'The group hours of {self.date_str} is non-downloadable')
                    self.times_df = None
                    break
                
                self.times_df = new_and_imporved_group_hours_html_reader(filepath, in_and_out_times=True, verbosity=0)
                
                os.remove(filepath)
                break
            except:
                print('oh no we failed on this attemp')
                i += 1
                try: 
                    x.kill()
                except:
                    print('no x to kill')
                    
    def rename_df_to_sql_columns(self):
        self.times_df_orig = self.times_df.copy()
        
        self.times_df = self.times_df.rename(columns={'Name':'name', 
                                                      'Job Code':'jobcode',
                                                      'Cost Code': 'costcode',
                                                      'Hours': 'hours',
                                                      'Time In': 'timein', 
                                                      'Time Out': 'timeout'})
        
        self.df_renamed_toggle = True
        
     
    def insertToLiveTable(self, table_name, df):
        
        # get table metadata
        metadata = MetaData()
        table_metadata = Table(table_name, metadata, autoload_with=engine, schema='live')
        columns = table_metadata.columns.keys()
        # make sure df columns match
        
        # matched_columns = [i for i in set(df.columns) if i in columns]
        # unmatched_columns = [i for i in set(df.columns) if i not in matched_columns]
        
        # if len(matched_columns) == len(columns):
        #     print(f'columns in provided df are a match to live.{table_name}')
        # else:
            
            
        #     print(f'columns in provided df are a mistmatch to live.{table_name}')
        #     raise Exception('Columns are not enough of a match & I didnt build something to handle this yet!')
        
        
        
        
        
        # check size of table before clearing it
        print_count_results('live', table_name, self.engine, 'before truncating')
        
        #truncate table
        connection = self.engine.connect()
        connection.execute(f"TRUNCATE TABLE live.{table_name}")
        connection.close()
        
        # check size after truncating to double check it worked
        print_count_results('live', table_name, self.engine, 'after truncating')
        
        # insert results
        df.to_sql(table_name, self.engine, schema='live', if_exists='append', index=False)
        
        # check size after inserting 
        print_count_results('live', table_name, self.engine, 'after inserting')
        
        
        
        
    def callMergeProc(self, proc_name, table_name):
        
        print_count_results('dbo', table_name, self.engine, 'before merge proc')
        
        connection = self.engine.raw_connection()
        cursor = connection.cursor()
        cursor.execute(f"call dbo.{proc_name}()")
        connection.commit()
        connection.close()
        
        print_count_results('dbo', table_name, self.engine, 'after merge proc')
        
        print(f'merge proc {proc_name} succeeded')
        
     
    def check_job_costcodes(self):
        
        columns = ['jobcode','costcode']
        
        table_name = 'job_costcodes'
        
        proc_name = 'merge_job_costcodes'
        
        
        # check to make sure we have the right names
        if not self.df_renamed_toggle:
            self.rename_df_to_sql_columns()

        # make a local copy
        df = self.times_df.copy()

        # get df with only 
        self.job_costcodes = df[columns]
        
        # get combo of the job & costcodes
        self.combos = self.job_costcodes.groupby(columns).size()
        
        # goal is to get to the grouped values
        self.combos = self.combos.reset_index()
        
        # get rid of the groupby count column
        self.combos = self.combos.drop(columns=[0])
        
        # go to the live table & insert values
        self.insertToLiveTable(table_name, self.combos)
        
        # call our merge proc
        self.callMergeProc(proc_name, table_name)
        
        self.job_costcodes_checked = True
        
        
        
    def insertGroupHours(self):
        if self.mergeTodayAvailable == True:
            proc_name = 'merge_clocktimes_today'
            table_name = 'clocktimes_today'
        elif self.mergeYesterdayAvailable == True:
            proc_name = 'merge_clocktimes_yesterday'
            table_name = 'clocktimes_yesterday'
        elif self.mergeRemediationAvailable == True:
            proc_name = 'merge_clocktimes'
            table_name = 'clocktimes'
            self.times_df.loc[:,'remediationtype'] = self.remediationtype
        
        
        # check to make sure we have the right names
        if not self.df_renamed_toggle:
            self.rename_df_to_sql_columns()        
        
        
        if not self.job_costcodes_checked:
            raise CheckpointNotReachedException('job & cost codes not checked yet')
            
        self.get_timeout_to_sqlformat()
        
        # go to the live table & insert values
        self.insertToLiveTable(table_name, self.times_df)
        
        # call our merge proc
        self.callMergeProc(proc_name, table_name)
        
    def get_timeout_to_sqlformat(self):
        
        self.times_df['timeout'] = pd.to_datetime(self.times_df['timeout'], errors='coerce')
        self.times_df['isclockedin'] = self.times_df['timeout'].isna()
        
        
        
        

#%%

# def insertInstananeous(times_df):
    
#     connection = engine.raw_connection()
#     cursor = connection.cursor()
#     cursor.execute("call live.mergeclocktimes_fromstaging()")
#     connection.commit()
#     connection.close()
    
#     return None

# def insertRemediated(remediation_days=1):
    
#     now = datetime.datetime.now()
#     date_str = (now - datetime.timedelta(days=remediation_days)).strftime('%m/%d/%Y')
#     times_df = getDatesTimesDF(date_str)
    
#     if times_df is None:
#         print(f'The date {date_str} returned results that wont be processed...')
#         return None
    
    
#     table_name = 'clocktimes_remediation'
    
#     print_count_results('live', table_name, engine, 'before truncating')
#     connection = engine.connect()
#     connection.execute(f"TRUNCATE TABLE live.{table_name}")
#     connection.close()
#     print_count_results('live', table_name, engine, 'after truncating')
    
#     times_df = times_df.rename(columns={'Name':'name', 
#                                         'Job Code':'jobcode',
#                                         'Cost Code': 'costcode',
#                                         'Hours': 'hours',
#                                         'Time In': 'timein', 
#                                         'Time Out': 'timeout'})
    
#     times_df['timeout'] = pd.to_datetime(times_df['timeout'], errors='coerce')
#     times_df['isclockedin'] = times_df['timeout'].isna()
    
    
#     # NEED TO FIGURE OUT WHAT TO DO IF WE HAVE REMEDIATED HORUS THAT ARE CLOCKED IN #???
#     # times_df = times_df[~times_df['timeout'].isna()]
    
#     # times_df = times_df[~times_df['timein'].isna()]
    
#     times_df.loc[:,'remediationtype'] = remediation_days
    
#     # need to have a check for when times_df is empty --- it might not make it this far if that happens
    
#     times_df.to_sql(table_name, engine, schema='live', if_exists='append', index=False)
    
#     print_count_results('live', table_name, engine, 'after inserting')
        
    
#     print_count_results('dbo', 'clocktimes', engine, 'before merge proc') # should be 0
    
#     connection = engine.raw_connection()
#     cursor = connection.cursor()
#     cursor.execute("call dbo.merge_clocktimes()")
#     connection.commit()
#     connection.close()
    
    
#     print_count_results('live', table_name, engine, 'after merge proc') # should be 0
#     print_count_results('dbo', 'clocktimes', engine, 'after merge proc') # should be 0
    
#     return None


def get_a_bunch_thisisaoneoff():
    daysback = 0
    daysbacktoo = 365
    for i in range(daysback, daysbacktoo):
        date_str = (datetime.datetime.now() - datetime.timedelta(days=365 - i)).strftime('%m/%d/%Y')
        x = insertGroupHours(date_str)
        x.doStuff()