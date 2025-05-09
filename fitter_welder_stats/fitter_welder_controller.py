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



output_dir = Path().home() / 'documents' / 'FitterWelderStatsPDFReports'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    
    

#%% define recipients


# this is the actual email list 
state_recipients = {'TN':['cwilson@crystalsteel.net', 'awhitacre@crystalsteel.net',
                          'rrichard@crystalsteel.net','emohamed@crystalsteel.net'],
                    'MD':['cwilson@crystalsteel.net',  'mmishler@crystalsteel.net',
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
file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

today = datetime.datetime.now()
month = today.month - 1
month_name = datetime.datetime(today.year, month, 1).strftime('%B')
if month == 12:
    year = today.year - 1
else:
    year = today.year
    

aggregate_data = fitter_welder_stats_month(month, year, production)
#%% get the previous month's data
# go back one day from the month we are running
two_months = datetime.datetime(today.year, month, 1) - datetime.timedelta(days=1)
# now set it to be the first day of the month
two_months = datetime.datetime(two_months.year, two_months.month, 1)

aggregate_data_2mo = find_old_file(two_months.month, two_months.year)
if aggregate_data_2mo is None:
    aggregate_data_2mo = fitter_welder_stats_month(two_months.month, two_months.year, production)

three_months = two_months - datetime.timedelta(days=1)
three_months = datetime.datetime(three_months.year, three_months.month, 1)
aggregate_data_3mo = find_old_file(three_months.month, three_months.year)
if aggregate_data_3mo is None:
    aggregate_data_3mo = fitter_welder_stats_month(three_months.month, three_months.year, production)
    
past_agg_data = {2:aggregate_data_2mo,
                 3:aggregate_data_3mo}
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
