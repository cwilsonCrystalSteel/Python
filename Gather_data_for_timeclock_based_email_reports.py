# -*- coding: utf-8 -*-
"""
Created on Thu May  6 11:09:44 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('c://users//cwilson//documents//python//Weekly Shop Hours Project//')
sys.path.append('c://users//cwilson//documents//python//Attendance Project//')
import pandas as pd
import glob
import os
import datetime
from TimeClock_Group_Hours import download_group_hours
from TimeClock_Tools_Employee_Location import download_most_current_employee_location_csv
from Read_Group_hours_HTML import new_output_each_clock_entry_job_and_costcode
import json




def clean_up_this_gunk(times_df, ei):

    # rename the columns b/c it makes it easier to work with
    ei = ei.rename(columns={'<NUMBER>':'ID',
                     '<FIRSTNAME>':'First',
                     '<LASTNAME>':'Last',
                     '<LOCATION>':'Location',
                     '<CLASS>':'Shift',
                     '<SCHEDULEGROUP>':'Productive',
                     '<DEPARTMENT>':'Department'})
    
    # combine the name fields to a combined single field    
    ei['Name'] = ei['First'] + ' '+ ei['Last']  
    # get rid of any duplicate names & keep the last entry with the newest ID number
    ei = ei.loc[~ei['Name'].duplicated(keep='last')]
    # remove employees without a productive/nonproductive grouping
    ei_shop = ei[~ei['Productive'].isna()]
    # get only people that have productive in their schedule group
    ei_shop = ei_shop[ei_shop['Productive'].str.contains('PRODUCTIVE')]
    # get only productive employees
    ei_prod = ei_shop[~ei_shop['Productive'].str.contains('NON')]

    


    # get only the shop employees in times_df
    times_df = times_df[times_df['Name'].isin(ei_shop['Name'])]
    # get the employees that are in ei_shop but not in times_df
    absent = ei_shop[~ei_shop['Name'].isin(times_df['Name'])]
    # this is to prevent an error
    absent_copy = absent.copy()
    # create a shop name column based on the Productive column
    absent_copy.loc[:,'Shop'] = absent.loc[:,'Productive'].str[:2]
    # put it back to the original
    absent = absent_copy.copy()
    # drop the columns not needed
    absent = absent.drop(columns=['First','Last','Location','Shift'])
    # reset the index of times_df
    times_df = times_df.reset_index(drop=True)
    

    

    code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))
   
    hours = times_df[~times_df['Job Code'].isin(code_changes['Delete Job Codes'])]
    # join the state to the hours df
    hours = hours.join(ei.set_index('Name')['Productive'].astype(str).str[:2], on='Name')
    # rename
    hours = hours.rename(columns={'Productive':'Location'})
    # create the indirect df starting with indirect cost codes
    indirect = hours[hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # remove those indirect cost code items from the hours df now
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
    # three_digit_indirects = hours[hours['Job #'].astype(str).str.len() < 4]
    # get the job codes that are less than 1000
    three_digit_indirects = hours[hours['Job #'] < 1000]
    # append the 3digit jobs that are indirect into indirect df
    indirect = indirect.append(three_digit_indirects)
    # delete the 3digit jobs from hours
    hours = hours[~hours.index.isin(three_digit_indirects.index)]
    # append what is left of hours df to direct
    direct = direct.append(hours)
    # sort the direct dataframe
    direct = direct.sort_values(['Name','Job #','Cost Code'])
    # sort the indirect dataframe
    indirect = indirect.sort_values(['Name','Job #','Cost Code'])        


    direct['Is Direct'] = True
    indirect['Is Direct'] = False
    

    
    return {'Absent':absent, 
            'Employee Information':ei_shop, 
            'Clocks Dataframe':times_df,
            'Direct':direct,
            'Indirect':indirect}




def skip_timeclock_automated_retrieval(times_df_html_path, ei_csv_path):
    times_df = new_output_each_clock_entry_job_and_costcode(times_df_html_path)
    ei = pd.read_csv(ei_csv_path)
    
    return clean_up_this_gunk(times_df, ei)
    


def get_clock_times_html_downloaded(start_date, end_date, exclude_terminated=True, download_folder="C:\\users\\cwilson\\downloads\\"):

    today = datetime.datetime.today().date()
    
    count = 0
    while count < 4:
        try:
            count += 1
            downloadedSuccessful = download_group_hours(start_date, end_date, download_folder)
            if downloadedSuccessful is None:
                return None
            
            # Grab all HTML files in downloads
            list_of_htmls = glob.glob(download_folder +"*.html") # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            group_hours_html = [f for f in list_of_htmls if "Hours" in f]
            # Get the most recent file for that state
            latest_html = max(group_hours_html, key=os.path.getctime)
            # get the file creation time as a datetime
            latest_html_time = datetime.datetime.fromtimestamp(os.path.getctime(latest_html))
            # if the file was created today, then use it, if not then throw error            
            if latest_html_time.date() == today:
                print(latest_html)
                # times_df = output_each_clock_entry_job_and_costcode(latest_html)
                times_df = new_output_each_clock_entry_job_and_costcode(latest_html)
                # delete the html file
                os.remove(latest_html)                
            else:
                '2' + 2
                
            break
        except Exception as e:
            print('\n\nDownloading group hours failed: ')
            print(e)
            print('\n\n')
            pass
        
    return times_df


def get_ei_csv_downloaded(exclude_terminated, download_folder="C:\\users\\cwilson\\downloads\\"):
    
    today = datetime.datetime.today().date()
    # latest_csv = "c://users//cwilson//downloads//f_this.csv"
    count = 0
    while count < 6:
        try:
            count += 1
            download_most_current_employee_location_csv(download_folder, exclude_terminated)
            # Grab all HTML files in downloads
            list_of_csvs = glob.glob(download_folder +"*.csv") # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            employe_info_csvs = [f for f in list_of_csvs if "Employee Information" in f]
            # Get the most recent file for that state
            latest_csv = max(employe_info_csvs, key=os.path.getctime)
            # get the file creation time as a datetime
            latest_csv_time = datetime.datetime.fromtimestamp(os.path.getctime(latest_csv))
            # if the file was created today, then use it, if not then throw error
            if latest_csv_time.date() == today:
                print(latest_csv)
                ei = pd.read_csv(latest_csv)
                # Delete the Employee Information CSV file from the downloads folder
                os.remove(latest_csv)                
            else:
                2 + "2"
            break
        except Exception as e:
            print('Downloading employee information failed: ')
            print(e)
            pass
    
    return ei


def get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=True, download_folder="C:\\users\\cwilson\\downloads\\", ei=None):
    
    
    
    times_df = get_clock_times_html_downloaded(start_date, end_date, exclude_terminated, download_folder)
    
    if ei is None:
        ei = get_ei_csv_downloaded(exclude_terminated, download_folder)
    
    if times_df is None:
        return None
    else:
        
    
        return clean_up_this_gunk(times_df, ei)
    
    




























