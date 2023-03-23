# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:35:31 2023

@author: CWilson
"""

# this one will pull the timeclock data
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import pandas as pd
from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
import datetime

state = 'TN'
today = datetime.datetime.now()
today_str = today.strftime("%m/%d/%Y")

def get_timeclock_summary(start_dt, end_dt, state, basis=None):
    
    now = datetime.datetime.now()
    end_date = end_dt.strftime('%m/%d/%Y')
    start_date = start_dt.strftime('%m/%d/%Y')
            
    
    if basis == None:
        start_dt_loop = start_dt
        start_date_loop = start_dt_loop.strftime('%m/%d/%Y')
        print('Getting Timeclock for: {}'.format(start_date_loop))
        basis_orig = get_information_for_clock_based_email_reports(start_date_loop, start_date_loop, exclude_terminated=False, ei=None, in_and_out_times=True) 
        basis = basis_orig.copy()
        ei = basis['Employee Information']
        for i in range(1, (end_dt-start_dt).days):
            start_dt_loop = start_dt_loop + datetime.timedelta(days=1)
            if start_dt_loop.date() > now.date():
                continue
            start_date_loop = start_dt_loop.strftime('%m/%d/%Y')
            print('Getting Timeclock for: {}'.format(start_date_loop))
            basis_additional = get_information_for_clock_based_email_reports(start_date_loop, start_date_loop, exclude_terminated=False, ei=ei, in_and_out_times=True) 
            basis['Direct'] = basis['Direct'].append(basis_additional['Direct'], ignore_index=True)
            basis['Indirect'] = basis['Indirect'].append(basis_additional['Indirect'], ignore_index=True)
    
    
    
    
    
    
    # give them an extra couple of hours for early clock ins
    # i doubt anyone is going to start 2nd shift after 3 a.m.
    start_dt_filter = start_dt.replace(hour=3, minute=0)
    end_dt_filter = end_dt.replace(hour=4, minute=0)
    
    
    direct = basis['Direct']
    indirect = basis['Indirect']
    hours = direct.append(indirect, ignore_index=True)
    # convert time in to a datetime
    hours['Time In'] = pd.to_datetime(hours['Time In'], errors='coerce')
    # get rid of any that don't convert
    hours = hours[~hours['Time In'].isna()]
    # get only records that the clock in time is for that work day
    hours = hours[(hours['Time In'] >= start_dt_filter) & (hours['Time In'] <= end_dt_filter)]
    
    # hours = hours[hours['Location'] == state]
    
    ei = basis['Employee Information']
    # set the index to the employee name
    ei = ei.set_index('Name')
    # get all employees at that state
    ei = ei[ei['Productive'].str.contains(state)]
    # only get hours when there is an employee match
    hours = ei[['Productive','Shift']].join(hours.set_index('Name'), how='inner')
    
    hours_productive = hours[~hours['Productive'].str.contains('NON')]
    # hours_productive = hours
    
    num_employees = pd.unique(hours_productive.index).shape[0]
    num_direct = hours_productive[hours_productive['Is Direct']]['Hours'].sum().round(2)
    num_indirect = hours_productive[~hours_productive['Is Direct']]['Hours'].sum().round(2)
    
    return {'Number Employees':num_employees, 'Direct Hours':num_direct, 'Indirect Hours':num_indirect}
