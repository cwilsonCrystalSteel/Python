# -*- coding: utf-8 -*-
"""
Created on Sun Jun  5 09:40:01 2022

@author: CWilson
"""

import pandas as pd
import os
from pathlib import Path
import datetime
import numpy as np
import copy
from Automate_MDI import do_mdi, eva_vs_hpt
from High_Indirect_Hours_Email_Report import summarize_by_direct_indirect
from High_Indirect_Hours_Email_Report import output_absent_dict
from High_Indirect_Hours_Email_Report import email_absent_list
from High_Indirect_Hours_Email_Report import return_output_dictionary
from High_Indirect_Hours_Email_Report import email_sub80_results
from High_Indirect_Hours_Email_Report import email_sub2lots_results
from High_Indirect_Hours_Email_Report import email_mdi
from High_Indirect_Hours_Email_Report import email_eva_vs_hpt
from High_Indirect_Hours_Email_Report import emaIL_attendance_hours_report
from High_Indirect_Hours_Email_Report import email_delivery_calendar_changelog
from High_Indirect_Hours_Email_Report import email_error_message
from Attendance_Hours_Per_week_v2 import run_attendance_hours_report
from Speedo_Dashboard.Post_to_GoogleSheet import get_production_worksheet_production_sheet
from TimeClock.pullGroupHoursFromSQL import get_date_range_timesdf_controller
from TimeClock.functions_TimeclockForSpeedoDashboard import return_basis_new_direct_rules

#%%

try:
    production_worksheet_outpath = Path.home() / 'Documents' / 'speedo_dashboard' / 'production_worksheet.csv'
    df = get_production_worksheet_production_sheet(proper_headers=False)
    df.to_csv(production_worksheet_outpath, index=False)
except:
    print('could not get the production_worksheet')

#%%

# get today as a datetime
today = datetime.datetime.now()
# get yesterday as a datetime  
yesterday = today - datetime.timedelta(days=1)
# convert yesterday to a string format for my functions
yesterday_str = yesterday.strftime("%m/%d/%Y")


# this is the actual email list 
state_recipients = {'TN':['cwilson@crystalsteel.net', 'awhitacre@crystalsteel.net',
                          'jshade@crystalsteel.net','rhagins@crystalsteel.net'],
                    'MD':['cwilson@crystalsteel.net',  'mmishler@crystalsteel.net',
                          'jkeith@crystalsteel.net','jlaird@crystalsteel.net'],
                    'DE':['cwilson@crystalsteel.net',  'jrodriguez@crystalsteel.net'],
                    'EC':['cwilson@crystalsteel.net']}

eva_hpt_recipients = ['cwilson@crystalsteel.net','awhitacre@crystalsteel.net',
                      'nreed@crystalsteel.net','agagnon@crystalsteel.net',
                      'jlaird@crystalsteel.net']


''' #un-comment this when debugging 

state_recipients = {'TN':['cwilson@crystalsteel.net'],
                    'MD':['cwilson@crystalsteel.net'],
                    'DE':['cwilson@crystalsteel.net'],
                    'EC':['cwilson@crystalsteel.net'],}
eva_hpt_recipients = ['cwilson@crystalsteel.net']
'''

#%%

# get the data from TIMECLOCK then do basic transformations to it
# basis = get_information_for_clock_based_email_reports(yesterday_str, yesterday_str, exclude_terminated=True)

times_df = get_date_range_timesdf_controller(yesterday_str, yesterday_str)
# if we got no records for yesterday, lets run yesterday from the bat-ready py file
if not times_df.shape[0]:
    try:
        from TimeClock.insertGroupHoursToSQL import insertGroupHours
        source = 'Automated_Email_Reporting_Controller'
        download_folder = Path.home() / 'downloads' / 'GroupHours_Yesterday'
        x = insertGroupHours(date_str=yesterday_str, source=source, download_folder=download_folder)
        x.doStuff()
        times_df = get_date_range_timesdf_controller(yesterday_str, yesterday_str)
    except Exception as e:
        email_error_message([f'Error {e}', 
                             f'Source: {source}',
                             f'yesterday_str: {yesterday_str}',
                             'There was no times_df available for {yesterday_str}, and then we failed to obtain a new timeclock record!'])

basis = return_basis_new_direct_rules(times_df)

# gets all of the absent shop employees who did not clock in by comparing 
# employee information list to names in the clocked in list
absent = basis['Absent']
# only gets the employees that have a schedule group of XX Productive
ei = basis['Employee Information']
ei['Location'] = ei['Productive'].str[:2]
# get rid of the non productive employees in ei
ei = ei[~ei['Productive'].str.contains('NON')]
# get the dataframe version of the time clock - 
clock_raw_df = basis['Clocks Dataframe']
# reset the clock_raw_df index back to numbers
clock_raw_df = clock_raw_df.reset_index(drop=True)


direct = basis['Direct']
indirect = basis['Indirect']




# group direct - sum hours, average the job #
direct_grouped = direct.groupby(by=['Name','Job Code','Cost Code', 'Job #']).agg({'Hours':'sum'})
# add a is direct tag
direct_grouped['Is Direct'] = True
# group the indirect
indirect_grouped = indirect.groupby(by=['Name','Job Code','Cost Code', 'Job #']).agg({'Hours':'sum'})
# add a is direct tag
indirect_grouped['Is Direct'] = False
# append the grouped direct & grouped indirect
clock_grouped_df = pd.concat([direct_grouped, indirect_grouped])
# clock_grouped_df = direct_grouped.append(indirect_grouped)
# reset the index but do not drop it
clock_grouped_df = clock_grouped_df.reset_index(drop=False)
# merge on the locaiton of the employee
clock_grouped_df = clock_grouped_df.merge(ei[['Name','Location']], on='Name', how='inner')
# replce no cost code with a space
clock_grouped_df['Cost Code'] = clock_grouped_df['Cost Code'].replace('no cost code','-')

# Summarize each employee based on # of direct/indirect hours
clock_summary_df = summarize_by_direct_indirect(clock_grouped_df)


#%%

lots_calendar_changelog_dir = Path(os.getcwd())/ 'Lots_schedule_calendar' / 'Change_Logs_v2'
if not os.path.exists(lots_calendar_changelog_dir):
    os.makedirs(lots_calendar_changelog_dir)
    
yesterday_lots_calendar_changelog = lots_calendar_changelog_dir / (yesterday.strftime('%Y-%m-%d') + '.csv')
if os.path.exists(yesterday_lots_calendar_changelog):
    print('calendar updates')
    try:
        email_delivery_calendar_changelog(yesterday_str, 
                                          file_name = yesterday_lots_calendar_changelog, 
                                          recipient_list = ['cwilson@crystalsteel.net'])
    except:
        print('email_delivery_calendar_changelog failed')
else:
    print(f'Trying to update calendar changes but no file found {yesterday_lots_calendar_changelog}')


# only run if yesterday was not sunday
if yesterday.weekday() != 6:

    
    output_absent = output_absent_dict(absent)
    
    for state in output_absent.keys():
        if not state in ['MD','DE','TN']:
            continue
        
        if output_absent:
            # assign the recipients
            output_absent[state]['Recipients'] = state_recipients[state]
            try:
                # call the email function for the absent employees
                email_absent_list(yesterday_str, state, copy.deepcopy(output_absent[state]))
            except Exception as e:
                email_error_message([f'Error {e}', 
                                     'email_absent_list',
                                     f'yesterday_str: {yesterday_str}',
                                     f'state: {state}'])
    
    
    
    # get the employees with less then 90% direct hours
    sub_80 = clock_summary_df[clock_summary_df['% Direct'] < 0.9]
    # transform to HTML tables for email friendly viewing
    output_sub80direct = return_output_dictionary(sub_80, clock_grouped_df)
    
    
    for state in output_sub80direct.keys():
        if not state in ['MD','DE','TN']:
            continue
        # assign the recipients    
        if output_sub80direct:
            output_sub80direct[state]['Recipients'] = state_recipients[state]

            try:
                # call the email function for the employees with less then 80% direct
                email_sub80_results(yesterday_str, state, copy.deepcopy(output_sub80direct[state]))
            except Exception as e:
                email_error_message([f'Error {e}', 
                                     'email_sub80_results',
                                     f'yesterday_str: {yesterday_str}',
                                     f'state: {state}'])
        
    
    

    
    
    # Get the employees with <= 1 lot clocked in to
    sub_2_lots = clock_summary_df[clock_summary_df['# Lots'] <= 1]
    # transform to HTML tables for email friendly viewing
    output_sub2lots = return_output_dictionary(sub_2_lots, clock_grouped_df)
    # email out each state's list of employees with <= 1 lot
 
    for state in output_sub2lots.keys():
        if not state in ['MD','DE','TN']:
            continue
        if output_sub2lots:
            # assign the recipients
            output_sub2lots[state]['Recipients'] = state_recipients[state]

            try:
                # call the email function for the employees with <= 1 lot
                email_sub2lots_results(date_str=yesterday_str, state=state, state_dict=copy.deepcopy(output_sub2lots[state]))
            except Exception as e:
                email_error_message([f'Error {e}', 
                                     'email_sub2lots_results',
                                     f'yesterday_str: {yesterday_str}',
                                     f'state: {state}'])
        

    
        
    # This one sends out the MDI stuff
    for state in state_recipients.keys():
        if not state in ['MD','DE','TN']:
            continue
        if basis['Direct'].shape[0]:
            # calculate the MDI stuff from the do_mdi function
            mdi_dict = do_mdi(basis, state, yesterday_str)
            # we get a None if there was nothing in fablisting for the date = yesterday_str
            if mdi_dict is None:
                email_error_message([f'mdi_dict is None for {state}',
                                     'email_mdi',
                                     f'yesterday_str: {yesterday_str}'])
                continue

            try:
                # email out the mdi email
                # this funtion also adds rrichard@crystalsteel.net & emohamed@crystalsteel.net 
                email_mdi(date_str=yesterday_str, state=state, state_dict=copy.deepcopy(mdi_dict), email_dict=state_recipients)
        
            except Exception as e:
                email_error_message([f'Error {e}', 
                                     'email_mdi',
                                     f'yesterday_str: {yesterday_str}',
                                     f'state: {state}'])
    
# only send EVA once a week - so on sundays        
# else:
    
    try:
        day_pcs = eva_vs_hpt(start_date=yesterday_str, end_date=yesterday_str)
        
        ten_days = (yesterday - datetime.timedelta(days=10)).strftime("%m/%d/%Y")
        
        ten_day_lot = eva_vs_hpt(start_date=ten_days, end_date=yesterday_str)
        
        sixty_days = (yesterday - datetime.timedelta(days=60)).strftime("%m/%d/%Y")
        
        sixty_day_job = eva_vs_hpt(start_date=sixty_days, end_date=yesterday_str)
        
        # change it so you only run the sixty day timespan & then just portion out the 10 day & yesterday
        
        eva_vs_hpt_dict = {'Yesterday':day_pcs,
                           '10 day':ten_day_lot,
                           '60 day':sixty_day_job}
            
        email_eva_vs_hpt(date_str=yesterday_str, eva_vs_hpt_dict=eva_vs_hpt_dict, email_recipients=eva_hpt_recipients)
        
    except Exception as e:
        print('could not send the email eva_vs_hpt_dict')
        print(e)
        email_error_message([f'Error {e}', 
                             'email_eva_vs_hpt',
                             f'day_pcs: {yesterday_str} to {yesterday_str}',
                             f'ten_day_lot: {ten_days} to {yesterday_str}'
                             f'sixty_day_job: {sixty_days} to {yesterday_str}'])

# run on fridays?
if yesterday.weekday() == 3:
    print('Run weekly attendance hours report here')
    
    for state in state_recipients.keys():
        if not state in ['MD','DE','TN']:
            continue
        try:
            attendance_hours = run_attendance_hours_report(state)
            filepath = attendance_hours['filepath']
            weekstart = attendance_hours['weekstart']
            
            emaIL_attendance_hours_report(weekstart, state, filepath, state_recipients)
        except Exception as e:
            print('could not send attendance hours report')
            print(e)
            email_error_message([f'Error {e}', 
                                 'run_attendance_hours_report / emaIL_attendance_hours_report',])
 #%%       
    
quit()

