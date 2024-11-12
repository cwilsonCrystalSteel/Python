# -*- coding: utf-8 -*-
"""
Created on Sun Jun  5 09:40:01 2022

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('c://users//cwilson//documents//python//Weekly Shop Hours Project//')
sys.path.append('c://users//cwilson//documents//python//Attendance Project//')
sys.path.append('c://users//cwilson//documents//python//Lots_schedule_calendar//')
sys.path.append('c://users//cwilson//documents//python//TimeClock//')
import pandas as pd
import os
import datetime
import numpy as np
import copy
# from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
from TEMPORARY_Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports

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
from Attendance_Hours_Per_week_v2 import run_attendance_hours_report
from Post_to_GoogleSheet import get_production_worksheet_production_sheet
from pullGroupHoursFromSQL import get_date_range_timesdf_controller
from functions_TimeclockForSpeedoDashboard import return_information_on_clock_data
#%%

try:
    df = get_production_worksheet_production_sheet(proper_headers=False)
    df.to_csv('c:\\users\\cwilson\\\documents\\python\\speedo_dashboard\\production_worksheet.csv', index=False)
except:
    print('could not get the production_worksheet')



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


#''' un-comment this when debugging 

state_recipients = {'TN':['cwilson@crystalsteel.net'],
                    'MD':['cwilson@crystalsteel.net'],
                    'DE':['cwilson@crystalsteel.net'],
                    'EC':['cwilson@crystalsteel.net'],}
eva_hpt_recipients = ['cwilson@crystalsteel.net']

#'''
#%%

# get the data from TIMECLOCK then do basic transformations to it
# basis = get_information_for_clock_based_email_reports(yesterday_str, yesterday_str, exclude_terminated=True)

times_df = get_date_range_timesdf_controller(yesterday_str, yesterday_str)
basis = return_information_on_clock_data(times_df)

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

direct = basis['Direct']
# direct['Is Direct'] = True
indirect = basis['Indirect']
# indirect['Is Direct'] = False

# # replace 'no cost code' to be empty
# clock_raw_df['Cost Code'] = clock_raw_df['Cost Code'].replace('no cost code', '')
# # create a series thats a combination of job & cost code
# clock_raw_df['Job CostCode'] = clock_raw_df['Job #'].astype(str) + '<>' + clock_raw_df['Cost Code']
# # add the productive column based on employees name
# clock_raw_df = clock_raw_df.merge(ei[['Name','Location']], on='Name', how='left')



# reset the clock_raw_df index back to numbers
clock_raw_df = clock_raw_df.reset_index(drop=True)



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


lots_calendar_changelog_dir = 'C:\\Users\\cwilson\\Documents\\Python\\Lots_schedule_calendar\\Change_Logs_v2\\'
yesterday_lots_calendar_changelog = lots_calendar_changelog_dir + yesterday.strftime('%Y-%m-%d') + '.csv'
if os.path.exists(yesterday_lots_calendar_changelog):
    print('calendar updates')
    try:
        email_delivery_calendar_changelog(yesterday_str, 
                                          file_name = yesterday_lots_calendar_changelog, 
                                          recipient_list = ['cwilson@crystalsteel.net', 'jturner@crystalsteel.net'])
    except:
        print('email_delivery_calendar_changelog failed')
else:
    print(f'Trying to update calendar changes but no file found {yesterday_lots_calendar_changelog}')


# only run if yesterday was not sunday
if yesterday.weekday() != 6:

    
    output_absent = output_absent_dict(absent)
    
    for state in output_absent.keys():
        
        if output_absent:
            # assign the recipients
            output_absent[state]['Recipients'] = state_recipients[state]
            # call the email function for the absent employees
            email_absent_list(yesterday_str, state, copy.deepcopy(output_absent[state]))
    
    
    
    # get the employees with less then 90% direct hours
    sub_80 = clock_summary_df[clock_summary_df['% Direct'] < 0.9]
    # transform to HTML tables for email friendly viewing
    output_sub80direct = return_output_dictionary(sub_80, clock_grouped_df)
    
    
    for state in output_sub80direct.keys():
        # assign the recipients    
        if output_sub80direct:
            output_sub80direct[state]['Recipients'] = state_recipients[state]
            # call the email function for the employees with less then 80% direct
            email_sub80_results(yesterday_str, state, copy.deepcopy(output_sub80direct[state]))
        
    
    

    
    
    # Get the employees with <= 1 lot clocked in to
    sub_2_lots = clock_summary_df[clock_summary_df['# Lots'] <= 1]
    # transform to HTML tables for email friendly viewing
    output_sub2lots = return_output_dictionary(sub_2_lots, clock_grouped_df)
    # email out each state's list of employees with <= 1 lot
 
    for state in output_sub2lots.keys():
        if output_sub2lots:
            # assign the recipients
            output_sub2lots[state]['Recipients'] = state_recipients[state]
            # call the email function for the employees with <= 1 lot
            email_sub2lots_results(yesterday_str, state, copy.deepcopy(output_sub2lots[state]))
        

    
        
    # This one sends out the MDI stuff
    for state in [i for i in state_recipients.keys() if i != 'EC']:
        if basis['Direct'].shape[0]:
            # calculate the MDI stuff from the do_mdi function
            mdi_dict = do_mdi(basis, state, yesterday_str)
            # email out the mdi email
            # this funtion also adds rrichard@crystalsteel.net & emohamed@crystalsteel.net 
            email_mdi(yesterday_str, state, copy.deepcopy(mdi_dict), state_recipients)
    
# only send EVA once a week - so on sundays        
# else:
    
    try:
        day_pcs = eva_vs_hpt(yesterday_str, yesterday_str)
        
        ten_days = (yesterday - datetime.timedelta(days=10)).strftime("%m/%d/%Y")
        
        ten_day_lot = eva_vs_hpt(ten_days, yesterday_str)
        
        sixty_days = (yesterday - datetime.timedelta(days=60)).strftime("%m/%d/%Y")
        
        sixty_day_job = eva_vs_hpt(sixty_days, yesterday_str)
        
        # change it so you only run the sixty day timespan & then just portion out the 10 day & yesterday
        
        eva_vs_hpt_dict = {'Yesterday':day_pcs,
                           '10 day':ten_day_lot,
                           '60 day':sixty_day_job}
            
        email_eva_vs_hpt(yesterday_str, eva_vs_hpt_dict, eva_hpt_recipients)
        
    except Exception as e:
        print('could not send the email eva_vs_hpt_dict')
        print(e)

# # run on fridays?
# if yesterday.weekday() == 3:
#     print('Run weekly attendance hours report here')
    
#     for state in state_recipients.keys():
#         attendance_hours = run_attendance_hours_report(state)
#         filepath = attendance_hours['filepath']
#         weekstart = attendance_hours['weekstart']
        
        
#         # filepath = 'C:\\Users\\cwilson\\Documents\\Productive_Employees_Hours_Worked_Report\\week_by_week_hours_of_employees ' + state + '.xlsx'
#         # weekstart = '05/22/2022'
        
        
#         emaIL_attendance_hours_report(weekstart, state, filepath, state_recipients)

 #%%       
    
quit()

