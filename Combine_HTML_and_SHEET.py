# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 13:13:01 2021

@author: CWilson
"""
import pandas as pd
from datetime import datetime
from TimeClock_Job_Code_Summary import get_most_current_html
from Read_TimeClock_HTML import return_direct_indirect_sums_dfs
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours

def return_combined_hours_weights(state = 'TN', start_date="03/06/1997", end_date="03/06/1997"):

    allowable_attempts = 5
    count = 0
    while count <= allowable_attempts:
        try:
            folder = "C:\\Users\\cwilson\\Documents\\Python\\Publish to Live\\temp_HTML_downloads\\"
            # Get the most up to date version of the hours HTML file downloaded
            get_most_current_html(state, start_date, end_date, folder)
            # Read the hours and return the dataframe 
            hours_list = return_direct_indirect_sums_dfs(state, folder)
            break
        except:
            count += 1
    
    # pull the direct hours df from the list of dfs
    direct = hours_list[0].copy()
 
    
    if state == 'TN':
        sheet_name1 = 'CSM QC Form'
    elif state == 'DE':
        sheet_name1 = 'CSF QC Form'
    elif state == 'MD':
        sheet_name1 = 'FED QC Form'
    
    
    # Grab the pieces produced

    gs = grab_google_sheet(sheet_name1, start_date, end_date)
    ''' This is because DailyFabListing has a day defined as 6am to 6am the next day
        I don't want to appply the 6am start_dt to the grab_google_sheet function
        because that is used in other scripts so it easiest to do it here '''
    # convert the start_date to a datetime
    start_dt = datetime.strptime(start_date, "%m/%d/%Y")
    # change the start time to 6 am
    start_dt = start_dt.replace(hour=5, minute=0, second=0)
    # remove anything in the gs dataframe that came before the start_dt
    gs = gs[gs['Timestamp'] >= start_dt]
    
    # get the earned hours that come from the models
    gs = apply_model_hours(gs)
    
    # Pull the job names out
    jobs = direct['Job Code'].copy().tolist()
    # Only pull the job number out
    jobs_in_direct = []
    for job in jobs:
        job = int(job[:4])
        jobs_in_direct.append(job)
    # Replace the job names with the job numbers
    direct['Job Code'] = jobs_in_direct
    
    # Find the jobs from the google sheet
    jobs_in_gs = pd.unique(gs['Job #']).tolist()
    # Combine the jobs from the google sheet and direct hours to one series
    all_jobs = list(set(jobs_in_gs + jobs_in_direct))
    # Initilize the output dataframe with the jobs as column headers
    output_df = pd.DataFrame(columns=all_jobs, index=['Weight', 'Hours', 'EVA Hours'])
    
    
    for job in jobs_in_gs:
        this_df = gs[gs['Job #'] == job]
        weight = this_df['Weight'].sum()
        eva_hours = this_df['Earned Hours'].sum()
        output_df.loc['Weight', job] = weight
        output_df.loc['EVA Hours', job] = eva_hours.round(2)
        
    for job in jobs_in_direct:
        hours = direct[direct['Job Code'] == job]['Total'].values[0]
        output_df.loc['Hours', job] = hours
        
    
    output_df = output_df.fillna(value=0)
    
    day = datetime.now().strftime("%m/%d/%Y")
    time = datetime.now().strftime("%I:%M %p")
    
    return [output_df, day, time]
