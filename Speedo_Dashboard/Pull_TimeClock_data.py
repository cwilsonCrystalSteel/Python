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

def get_timeclock_summary(today_str, state, basis=None):
    state = 'TN'
    # today = datetime.datetime.now()
    # today_str = today.strftime("%m/%d/%Y")
    
    
    today = datetime.datetime.strptime(today_str, "%m/%d/%Y")
    if basis == None:
        basis = get_information_for_clock_based_email_reports(today_str, today_str, exclude_terminated=False, ei=None) 
    
    direct = basis['Direct']
    indirect = basis['Indirect']
    hours = direct.append(indirect, ignore_index=True)
    hours = hours[hours['Location'] == state]
    
    ei = basis['Employee Information']
    # set the index to the employee name
    ei = ei.set_index('Name')
    # get all employees at that state
    ei = ei[ei['Productive'].str.contains(state)]
    
    hours_productive = ei[['Productive','Shift']].join(hours.set_index('Name'), how='inner')
    
    hours_productive = hours_productive[~hours_productive['Productive'].str.contains('NON')]
    
    num_employees = pd.unique(hours_productive['Name']).shape[0]
    num_direct = hours_productive[hours_productive['Is Direct']]['Hours'].sum().round(2)
    num_indirect = hours_productive[~hours_productive['Is Direct']]['Hours'].sum().round(2)
    
    return {'employees':num_employees, 'direct':num_direct, 'indirect':num_indirect}