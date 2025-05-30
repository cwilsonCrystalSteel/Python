# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:37:10 2023

@author: CWilson
"""

from Speedo_Dashboard.Pull_Fablisting_data import get_fablisting_plus_model_summary
from Speedo_Dashboard.Pull_TimeClock_data import get_timeclock_summary
from Speedo_Dashboard.Post_to_GoogleSheet import post_observation, move_to_archive, get_jobs_to_exclude
import datetime
from pathlib import Path
import os

json_path = Path(os.getcwd()) / 'production-dashboard-other-e051ae12d1ef.json'

weekly_sheet_info_dict = {'sheet_key':'1RZKV2-jt5YOFJNKM8EJMnmAmgRM1LnA9R2-Yws2XQEs',
                          'json_file':json_path
                          }
'''
# Variables for easier testing
state = 'TN'
sheet = 'CSM QC Form' 
'''
# get current time
now = datetime.datetime.now()
# set the time to 6 am
now_dt = now.replace(hour=6, minute=0, second=0, microsecond=0)
# get start_dt as most recent sunday
start_dt = now_dt - datetime.timedelta(days=now_dt.weekday()+1)
# set end_dt to saturday
end_dt = start_dt + datetime.timedelta(days=6)
# change time of end_dt to midnight on saturday
end_dt = end_dt.replace(hour=23, minute=59)
print('running the speedo dashboard for {} to {}'.format(start_dt, end_dt))

try:
    
    exclude_jobs_dict = get_jobs_to_exclude(weekly_sheet_info_dict)
    exclude_jobs_dict['TN'] = exclude_jobs_dict['CSM']
    exclude_jobs_dict['DE'] = exclude_jobs_dict['CSF']
    exclude_jobs_dict['MD'] = exclude_jobs_dict['FED']
except:
    exclude_jobs_list = [3122,2214]
    exclude_jobs_dict = {'TN':exclude_jobs_list, 'MD':exclude_jobs_list, 'DE':exclude_jobs_list, 
                         'CSM':exclude_jobs_list, 'FED':exclude_jobs_list, 'CSF':exclude_jobs_list}



# get the results of each states hours - a dict divied up by state
timeclock_summary = get_timeclock_summary(start_dt, end_dt, states=None, basis=None, output_productive_report=True, exclude_jobs_dict=exclude_jobs_dict)
# timeclock_summary2 = get_timeclock_summary(start_dt, end_dt, states=None, basis=timeclock_summary['basis'], output_productive_report=False, exclude_jobs_list=exclude_jobs_list)

fablisting_summary = {}
for sheet in ['CSM QC Form','CSF QC Form','FED QC Form']:
    print(sheet)
    fablisting_summary[sheet] = get_fablisting_plus_model_summary(start_dt, end_dt, sheet=sheet, exclude_jobs_dict=exclude_jobs_dict, output_fablisting_copy=True)


for state in ['TN','MD','DE']:
    
    try:
        if state == 'TN':
           sheet = 'CSM QC Form'        
        elif state == 'DE':
           sheet = 'CSF QC Form'
        elif state == 'MD':
           sheet = 'FED QC Form'
                
        print(state, sheet, 'updating google sheet')
        
        run_for_date = start_dt.strftime('%m/%d/%Y')
        gsheet_dict = {'Date':run_for_date}
        gsheet_dict.update(fablisting_summary[sheet])
        gsheet_dict.update(timeclock_summary[state])
        
        post_observation(gsheet_dict, sheet_name=sheet[:3], google_sheet_info_dict=weekly_sheet_info_dict)
        
        move_to_archive(shop=sheet[:3], google_sheet_info_dict=weekly_sheet_info_dict, dashboard_name = 'weekly')

    except Exception as e:
        print(state)
        print('\n\n')
        print(e)
        print('\n\n')
        continue

# if gsheet_dict['Direct Hours']:
#     post_observation(gsheet_dict)
# else:
#     print('There was no direct horus so I wont post anything')

predictor = None


# if predictor != None:
#     predictor = get_prediction_formula()
#     post_predictor()


# if now.hour == 6:
#     ''' get the previous days shit for the summary tab '''
#     start_dt = now - datetime.timedelta(days=1)
#     start_dt = start_dt.replace(hour=6, minute=0, second=0, microsecond=0)
#     end_dt = start_dt + datetime.timedelta(days=1)
#     fablisting_summary = get_fablisting_plus_model_summary(start_dt, end_dt, sheet=sheet)
#     timeclock_summary = get_timeclock_summary(start_dt, end_dt, state=state, basis=None)
    
#     run_for_date = start_dt.strftime('%m/%d/%Y')
#     gsheet_dict = {'Date':run_for_date}
#     gsheet_dict.update(fablisting_summary)
#     gsheet_dict.update(timeclock_summary)    
#     post_observation(gsheet_dict, isReal=True, sheet_name='Day Summary')

#quit()