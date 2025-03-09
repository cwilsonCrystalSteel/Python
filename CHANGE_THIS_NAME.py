# -*- coding: utf-8 -*-
"""
Created on Tue Jul 20 11:19:21 2021

@author: CWilson
"""

import pandas as pd
from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
import json
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2
import os
from pathlib import Path

download_folder = Path.home() / 'downloads'

def download_data(start_date, end_date, download_folder=download_folder):
    basis = get_information_for_clock_based_email_reports(start_date, 
                                                          end_date,
                                                          exclude_terminated=False,  
                                                          download_folder=download_folder)
    return basis


def get_production_dashboard_data(start_date, end_date, base_data):
    
    #%% getting the hours here 
    
    ''' Change this to just use direct=basis['Direct'] & same for indirect '''
    code_changes_filepath = Path(os.getcwd()) / 'job_and_cost_code_changes.json'
    # load up the code changes file from the json
    code_changes = json.load(open(code_changes_filepath))
    # get the clocks dataframe
    hours = base_data['Clocks Dataframe']
    # get rid of any troublesome clocks bc that is easiest
    hours = hours[~hours['Job #'].isna()]
    # get the employee information out
    ei = base_data['Employee Information']
    # get the productive column from the employee information onto the hours df based on employee name
    hours = hours.join(ei.set_index('Name')['Productive'], on='Name', how='right')
    # get the location from the productive column 
    hours['State'] = hours['Productive'].str[:2]
    # delete the productive column & start & end times
    hours = hours.drop(columns=['Productive'])
    # get rid of the job codes that need to be deleted
    hours = hours[~hours['Job Code'].isin(code_changes['Delete Job Codes'])]
    # create the indirect df starting with indirect cost codes
    indirect = hours[hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # remove those from the hours df now
    hours = hours[~hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # append the indirect job codes
    indirect = indirect.append(hours[hours['Job Code'].isin(code_changes['Indirect Job Codes'])])
    # remove those from the hours df now
    hours = hours[~hours['Job Code'].isin(code_changes['Indirect Job Codes'])]
    # start with the things that appear to be indirect but are actually direct
    direct = hours[hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # remove those from the hours df now
    hours = hours[~hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # get the 3 digit items from hours df and move to indirect
    three_digit_indirects = hours[hours['Job #'] < 1000]
    # append the 3digit jobs that are indirect into indirect df
    indirect = indirect.append(three_digit_indirects)
    # delete the 3digit jobs from hours
    hours = hours[~hours.index.isin(three_digit_indirects.index)]
    # append what is left of hours df to direct
    direct = direct.append(hours)
    
    # direct = base_data['Direct']
    # indirect = base_data['Indirect']
    
    # group the hours by state, then by job # 
    grouped_hours = direct.groupby(['State','Job #']).sum()
    # change the index
    grouped_hours = grouped_hours.reset_index(drop=False)
    # change around the index
    grouped_hours = grouped_hours.set_index('Job #')
    # change the column name from hours to worked hours
    grouped_hours = grouped_hours.rename(columns={'Hours':'Worked Hours'})
    
    #%% creating the output dataframe
    
    states = ['TN','MD','DE']
    output_dict = {}
    
    for state in states:
        # get the fablisting sheet names
        if state == 'TN':
            sheet_name = 'CSM QC Form'
        elif state == 'MD':
            sheet_name = 'FED QC Form' 
        elif state == 'DE':
            sheet_name = 'CSF QC Form'
        
        # get the fablisting for the date range
        fablisting = grab_google_sheet(sheet_name, start_date, end_date, start_hour=5)
        # apply the model hours to the fablisting dataframe
        fablisting = apply_model_hours(fablisting_df = fablisting, 
                                       how = 'old way', 
                                       shop = sheet_name[:3])
        
        # only keep these columns
        fablisting = fablisting[['Piece Mark - REV','Job #','Lot #','Weight','Earned Hours']]
        # create the dataframe by grouping fablisting by the job # & summing
        fablisting_grouped = fablisting.groupby(['Job #']).sum().reset_index()
        # get only that states hours
        states_hours = grouped_hours[grouped_hours['State'] == state]
        # get rid of the state column and rest the index
        states_hours = states_hours['Worked Hours'].reset_index()
        # merge the dataframes 
        output_df = fablisting_grouped.merge(states_hours, how='outer')
        # replace nan values with zeros
        output_df = output_df.fillna(0)
        # set the index to be the job #
        output_df = output_df.set_index('Job #')
        # sort the dataframe by the Earned hours, then weight, then worked hours
        output_df = output_df.sort_values(['Earned Hours','Weight','Worked Hours'], ascending=False)
        # reorder the columns bc it was easiest to make the change here
        output_df = output_df[['Weight','Worked Hours','Earned Hours']]
        # set the output_df into that states part of the output dict
        output_dict[state] = output_df
    
    
    
    return output_dict



















