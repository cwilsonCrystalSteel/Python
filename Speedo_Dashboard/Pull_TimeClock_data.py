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
import json
code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))


state = 'TN'
today = datetime.datetime.now()
today_str = today.strftime("%m/%d/%Y")
basis = get_information_for_clock_based_email_reports(today_str, today_str, exclude_terminated=False, ei=None) 

direct = basis['Direct']
indirect = basis['Indirect']


ei = basis['Employee Information']
# set the index to the employee name
ei = ei.set_index('Name')
# get all employees at that state
ei = ei[ei['Productive'].str.contains(state)]

num_employees = pd.unique(direct.append(indirect, ignore_index=True)['Name'])