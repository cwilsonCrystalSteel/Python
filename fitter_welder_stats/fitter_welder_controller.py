# -*- coding: utf-8 -*-
"""
Created on Mon Apr 21 16:05:04 2025

@author: Netadmin
"""

import datetime
from pathlib import Path
import os
from fitter_welder_stats.Fitter_Welder_Stats_v2 import fitter_welder_stats_month, find_old_file
from fitter_welder_stats.Fitter_Welder_stats_PDF_report import pdf_report
from fitter_welder_stats.Fitter_Welder_Stats_emailing import email_pdf_report, email_error_message
import pandas as pd


file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
output_dir = Path().home() / 'documents' / 'FitterWelderStatsPDFReports'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    
    

#%% define recipients


# this is the actual email list 
state_recipients = {'TN':['cwilson@crystalsteel.net', 'awhitacre@crystalsteel.net',
                          'ewilson@crystalsteel.net',
                          'rrichard@crystalsteel.net','emohamed@crystalsteel.net'],
                    'MD':['cwilson@crystalsteel.net',  'mmishler@crystalsteel.net',
                          'jkeith@crystalsteel.net',
                          'rrichard@crystalsteel.net','emohamed@crystalsteel.net'],
                    'DE':['cwilson@crystalsteel.net',  'jrodriguez@crystalsteel.net',
                          'rrichard@crystalsteel.net','emohamed@crystalsteel.net']
                    }

production=True
''' #un-comment this when debugging 

state_recipients = {'TN':['cwilson@crystalsteel.net'],
                    'MD':['cwilson@crystalsteel.net'],
                    'DE':['cwilson@crystalsteel.net']
                    }
production = False

'''



#%% run the data 

# current date
today = datetime.datetime.now()
# go to the last day of the previous month
month_start = datetime.datetime(today.year, today.month, 1) - datetime.timedelta(days=1)
# now go back to first day of previous month
month_start = datetime.datetime(month_start.year, month_start.month, 1)
# grab its month number
month = month_start.month
# convert month to its calendar name
month_name = month_start.strftime('%B')
# get the year
year = month_start.year

    
# now get the data for that month
aggregate_data = fitter_welder_stats_month(month, year, production)

#%% get the previous month's data

number_months = 12

past_agg_data = {}
for i in range(1,1+number_months):
    # reset the value of months_ago to one month before month_start
    months_ago = datetime.datetime(month_start.year, month_start.month, 1) - datetime.timedelta(days=1) 
    months_ago = datetime.datetime(months_ago.year, months_ago.month, 1)
    # loop thru to iteratively subtract one month at a time 
    for j in range(1,i):
        months_ago = datetime.datetime(months_ago.year, months_ago.month, 1) - datetime.timedelta(days=1) 
        months_ago = datetime.datetime(months_ago.year, months_ago.month, 1)
        
    aggregate_data_months_ago = find_old_file(months_ago.month, months_ago.year)
    if aggregate_data_months_ago is None:
        aggregate_data_months_ago = fitter_welder_stats_month(months_ago.month, months_ago.year, production)
        past_agg_data[i] = aggregate_data_months_ago['filepath']
    else:
        # if we get it from find_old_file- its a filepath
        past_agg_data[i] = aggregate_data_months_ago

''' Handle sending messages out about the missing data'''
#%% create pdfs & mail


    


for state in ['MD','DE','TN']:
    try:
        output_file = f"FitterWelderStats-{state}-{month_name}-{year}_{file_timestamp}.pdf"
        output_filepath = output_dir / output_file
        pdfreport = pdf_report(state, aggregate_data=aggregate_data, output_pdf=output_filepath, past_agg_data=past_agg_data)
        pdfreport.build_report()
    
    
        wrong_state = aggregate_data.get('wrong_state').get(state)
        didnt_work = aggregate_data.get('didnt_work').get(state)
        invalid_id = aggregate_data.get('invalid_id').get(state)
        
        # nice names for titles in email
        dfs_to_fix_dict = {'Employee Does not Work in this State':wrong_state, f'Employee Did not Work in {month_name}':didnt_work, 'Invalid ID':invalid_id}
        # limit to dfs > 0 and sort ascending
        dfs_to_fix_dict = {name:df for name,df in sorted(dfs_to_fix_dict.items(), key=lambda item: item[1].shape[0]) if df.shape[0]}
        
            
    
        email_pdf_report(pdf_filepath=output_filepath, 
                          month_name=month_name, 
                          year=year, 
                          dfs_to_fix_dict = dfs_to_fix_dict,
                          state=state, 
                          recipients=state_recipients[state])
    
    except Exception as e:
        print(e)
        email_error_message(['fitter_welder_controller.py',
                             str(e),
                             f"{state} {month_name} {year}"])
