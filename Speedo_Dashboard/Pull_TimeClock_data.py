# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:35:31 2023

@author: CWilson
"""

# this one will pull the timeclock data
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
# from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')

from TEMPORARY_Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
import datetime
import pandas as pd
import numpy as np

state = 'TN'
today = datetime.datetime.now()
today_str = today.strftime("%m/%d/%Y")

def get_timeclock_summary(start_dt, end_dt, states=None, basis=None, output_productive_report=False, exclude_jobs_dict=None):
    if states == None:
        states = ['TN','MD','DE']
        
        
    now = datetime.datetime.now()
    end_date = end_dt.strftime('%m/%d/%Y')
    start_date = start_dt.strftime('%m/%d/%Y')
            
    
    if basis == None:
        start_dt_loop = start_dt
        start_date_loop = start_dt_loop.strftime('%m/%d/%Y')
        print('Getting Timeclock for: {}'.format(start_date_loop))
        basis_orig = get_information_for_clock_based_email_reports(start_date_loop, start_date_loop, exclude_terminated=False, ei=None, in_and_out_times=True) 
        
        '''
        # hokie work around
        # on 5/28 there are no records at all so it was failing to create the times_df variable
        # when there was no times_df, using the basis_orig.copy() was failing
        # i am just setting it up to try and move one day forward
        # this will still fail when the start of the week has no records until there the next day is available
        # this sucks
        # i need to implement some way to pass an empty times_df so that we can have a zero hour thing
        #
        
        try:
            basis = basis_orig.copy()
        except:
            start_dt_loop = start_dt + datetime.timedelta(days=1)
            start_date_loop = start_dt_loop.strftime('%m/%d/%Y')
            basis_orig = get_information_for_clock_based_email_reports(start_date_loop, start_date_loop, exclude_terminated=False, ei=None, in_and_out_times=True) 
            basis = basis_orig.copy()
         '''   
        
        # on 7/24 I started having troubles with timeclock when no records where found.
        # it looks like the same issue as from 5/28, and I think it is
        # but the TimeClock_Group_hours was not failing correctly?
        # i'm going to loop for 2 days to try and move up the start date
        #
        if isinstance(basis_orig, bool):
            print('The type of basis_orig is a bool! This only happens when the call to basis_orig errors out')
            print('when I implemented this (2024-07-24), it was happeneing bc the timeclock was returning no records and still trying to click the disabled download button')
            for i in range(0,3):
                start_dt_loop += datetime.timedelta(days=1)
                if start_dt_loop >= end_dt:
                    # we are trying to get a start date that is after the provided end_dt!
                    print(f"Value of start_dt_loop {start_dt_loop} exceeds the passed end date {end_dt}")
                    break
                
                start_date_loop = start_dt_loop.strftime('%m/%d/%Y')
                print('Getting Timeclock for: {}'.format(start_date_loop))
                basis_orig = get_information_for_clock_based_email_reports(start_date_loop, start_date_loop, exclude_terminated=False, ei=None, in_and_out_times=True) 
    
                
                if isinstance(basis_orig, bool):
                    continue
                else:
                    # get out of the loop if we get a dict(?)
                    break
            
        try:
            basis = basis_orig.copy()
        except:
            print('Could not get the basis from basis_orig, because it was not a copyable type')
            
            
            
        ei = basis['Employee Information']
        for i in range(1, (end_dt-start_dt).days):
            start_dt_loop = start_dt_loop + datetime.timedelta(days=1)
            if start_dt_loop.date() > now.date():
                continue
            start_date_loop = start_dt_loop.strftime('%m/%d/%Y')
            print('Getting Timeclock for: {}'.format(start_date_loop))
            basis_additional = get_information_for_clock_based_email_reports(start_date_loop, start_date_loop, exclude_terminated=False, ei=ei, in_and_out_times=True)
            basis['Direct'] = pd.concat([basis['Direct'], basis_additional['Direct']])
            basis['Indirect'] = pd.concat([basis['Indirect'], basis_additional['Indirect']])
    
    
    
    

    
    # give them an extra couple of hours for early clock ins
    # i doubt anyone is going to start 2nd shift after 3 a.m.
    start_dt_filter = start_dt.replace(hour=3, minute=0)
    # end_dt_filter = end_dt.replace(hour=4, minute=0)
    end_dt_filter = end_dt + datetime.timedelta(hours=4)
    
    
    direct = basis['Direct']
    indirect = basis['Indirect']
    
 
    
    hours = pd.concat([direct, indirect])
    # hours = direct.append(indirect, ignore_index=True)
    # convert time in to a datetime
    hours['Time In'] = pd.to_datetime(hours['Time In'], errors='coerce')
    # get rid of any that don't convert
    hours = hours[~hours['Time In'].isna()]
    # get only records that the clock in time is for that work day
    hours = hours[(hours['Time In'] >= start_dt_filter) & (hours['Time In'] <= end_dt_filter)]
    # pull the employee information df out of the basis dict
    ei = basis['Employee Information']
    # set the index to the employee name
    ei = ei.set_index('Name')
    
       
    
    # init the output dict so that we can have a dict of dicts
    output = {}
    
    for state in states:
        output[state] = {}
        # get all employees at that state
        ei_state = ei[ei['Productive'].str.contains(state)]
        # only get hours when there is an employee match
        hours_state = ei_state[['Productive','Shift']].join(hours.set_index('Name'), how='inner')
        # get the jobs to exclude
        excluded_jobs = exclude_jobs_dict[state]
        # filter those jobs out
        hours_state = hours_state[~hours_state['Job #'].isin(excluded_jobs)]
        # get the hours that are by productive employees 
        hours_productive = hours_state[~hours_state['Productive'].str.contains('NON')]
        # count the number of productive employees
        num_employees = pd.unique(hours_productive.index).shape[0]
        # count number of direct hours
        num_direct = np.round(hours_productive[hours_productive['Is Direct']]['Hours'].sum(), 2)
        # count number of indirect hours
        num_indirect = np.round(hours_productive[~hours_productive['Is Direct']]['Hours'].sum(), 2)
        # put into the output dict
        output[state] = {'Number Employees':num_employees, 'Direct Hours':num_direct, 'Indirect Hours':num_indirect}
        
        try:
            if output_productive_report:
                group_like_timeclock_report_TNproductive = hours_productive.copy()
                group_like_timeclock_report_TNproductive['date'] = group_like_timeclock_report_TNproductive['Time In'].dt.date
                group_like_timeclock_report_TNproductive = group_like_timeclock_report_TNproductive.groupby(['Is Direct','Job Code','date']).sum()['Hours']
                group_like_timeclock_report_TNproductive.to_excel('c:\\users\\cwilson\\downloads\\report_like_TNproductive.xlsx')
                output[state]['productive_report'] = group_like_timeclock_report_TNproductive
        except:
            print('could not make TN productive like report')

    output['basis'] = basis
    return output
