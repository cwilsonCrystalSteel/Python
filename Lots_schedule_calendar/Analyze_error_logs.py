# -*- coding: utf-8 -*-
"""
Created on Fri Jun  3 07:46:06 2022

@author: CWilson
"""

import pandas as pd
import os
import datetime



def get_errors_from_logs(person_name=None, lots_log = False):
    
    error_dir = 'C:\\Users\\cwilson\\Documents\\Python\\Lots_schedule_calendar\\Error_Logs'
    
    today = datetime.datetime.now()
    
    time_buffer = today - datetime.timedelta(minutes=60)
    
    print('Showing errors that occured between now & {}'.format(time_buffer.strftime('%m/%d/%Y %H:%M')))
    
    error_logs = os.listdir(error_dir)
    
    if person_name != None:
        error_logs = [i for i in error_logs if person_name in i]
    
    
    error_logs = [i for i in error_logs if 'Shipping Schedule' in i]
    
    errors_df = pd.DataFrame()
    
    
    
    for log in error_logs:
        
        log_filename = error_dir + '\\' + log
        
        creation_time = os.path.getctime(log_filename)
        
        creation_dt = datetime.datetime.fromtimestamp(creation_time)
        # dont bother adding files that are older than two weeks 
        if creation_dt < time_buffer:
            continue
        
        log_df = pd.read_csv(log_filename, header=1)
        
        if log_df.shape[0] == 0:
            continue
        
        log_df['Error Type'] = log.split('-')[1].split('for')[0].strip()
        
        errors_df = errors_df.append(log_df, ignore_index=False)
        
    
    errors_df = errors_df.drop_duplicates()
    
    return errors_df


scott = get_errors_from_logs('Scott')

all_ = get_errors_from_logs()


