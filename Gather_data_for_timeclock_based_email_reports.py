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
from Read_Group_hours_HTML import new_output_each_clock_entry_job_and_costcode, new_and_imporved_group_hours_html_reader
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
    
    #new way of getting direct / indirect
    # 3 digit codes are indirect - except for recieving
    # 5 digit codes are direct - except for CAPEX (idk how to tell if it is CAPEX)
    
    # split the cost code on a space or a slash
    times_df['Job #'] = times_df['Cost Code'].str.split(r'\s|\\').str[0]
    # get the shop site by taking the first 2 characters of the PRODUCTIVE tag from the EI dataframe
    times_df = times_df.join(ei.set_index('Name')['Productive'].astype(str).str[:2], on='Name')
    # rename that to location
    times_df = times_df.rename(columns={'Productive':'Location'})
    
    ''' remove CSM shop B employees here '''
    # remove employee ids: 2001, 2015, 2029
    # get the names of employees 2001, 2015, 2029
    shop_b_employees = ei[ei['ID'].isin([2001,2015,2029,2242,2261,2007,2241])]
    # drop those names from times_df
    times_df = times_df[~times_df['Name'].isin(list(shop_b_employees['Name']))]
    ''' end of removing shop B employees '''
    
    # get items where length of the job # is 5 or the cost code says recieving (or I could do job #  = 250)
    direct = times_df[(times_df['Job #'].str.len() == 5) | (times_df['Cost Code'].str.contains('RECEIVING'))].copy()
    # inidirect is whatver is not in the direct dataframe (in this instance)
    indirect = times_df.loc[~times_df.index.isin(direct.index)].copy()
    # now to convert the job number to the actual job number - only need to remove the first digit from the 5 digit jobs
    # direct['Job #'] = direct['Job #'].str[-4:]
    # 2023-09-28: changing this to the 1st 4 digits b/c apparantly thats what it is 
    direct['Job #'] = direct['Job #'].str[:4]
    
    # code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))
    # #incorporate the ffect of the new jobcode/costcode style in timeclock
    # for category in code_changes.keys():
    #     # get the list of jobs/costcodes to be changed
    #     changes = code_changes[category]
    #     # strip out everything but the jobnumber - that is the new job code
    #     alternate_style = [i.split(' ')[0] for i in changes]
    #     # add the 2 lists together and set it back into the dict[key]
    #     code_changes[category] = changes + alternate_style
        
    # hours = times_df[~times_df['Job Code'].isin(code_changes['Delete Job Codes'])]
    # # join the state to the hours df
    # hours = hours.join(ei.set_index('Name')['Productive'].astype(str).str[:2], on='Name')
    # # rename
    # hours = hours.rename(columns={'Productive':'Location'})
    # # create the indirect df starting with indirect cost codes
    # indirect = hours[hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # # remove those indirect cost code items from the hours df now
    # hours = hours[~hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # # append the indirect job codes
    # indirect = indirect.append(hours[hours['Job Code'].isin(code_changes['Indirect Job Codes'])])
    # # remove those from the hours df now
    # hours = hours[~hours['Job Code'].isin(code_changes['Indirect Job Codes'])]
    # # start with the things that appear to be indirect but are actually direct
    # direct = hours[hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # # remove those from the hours df now
    # hours = hours[~hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # # get the 3 digit items from hours df and move to indirect
    # # three_digit_indirects = hours[hours['Job #'].astype(str).str.len() < 4]
    # # get the job codes that are less than 1000
    # three_digit_indirects = hours[hours['Job #'] < 1000]
    # # append the 3digit jobs that are indirect into indirect df
    # indirect = indirect.append(three_digit_indirects)
    # # delete the 3digit jobs from hours
    # hours = hours[~hours.index.isin(three_digit_indirects.index)]
    # # append what is left of hours df to direct
    # direct = direct.append(hours)
    # # sort the direct dataframe
    # direct = direct.sort_values(['Name','Job #','Cost Code'])
    # # sort the indirect dataframe
    # indirect = indirect.sort_values(['Name','Job #','Cost Code'])        


    direct['Is Direct'] = True
    indirect['Is Direct'] = False
    

    
    return {'Absent':absent, 
            'Employee Information':ei_shop, 
            'Clocks Dataframe':times_df,
            'Direct':direct,
            'Indirect':indirect}




def skip_timeclock_automated_retrieval(times_df_html_path, ei_csv_path, in_and_out_times=False):
    # times_df = new_output_each_clock_entry_job_and_costcode(times_df_html_path, in_and_out_times=in_and_out_times)
    try:
         times_df = new_and_imporved_group_hours_html_reader(times_df_html_path, in_and_out_times=in_and_out_times)
    except Exception:
         times_df = new_output_each_clock_entry_job_and_costcode(times_df_html_path, in_and_out_times=in_and_out_times)
    ei = pd.read_csv(ei_csv_path)
    
    return clean_up_this_gunk(times_df, ei)
    


def get_clock_times_html_downloaded(start_date, end_date, exclude_terminated=True, download_folder="C:\\users\\cwilson\\downloads\\", in_and_out_times=False):

    today = datetime.datetime.today().date()
    
    count = 0
    while count < 4:
        try:
            count += 1
            timeclocker = download_group_hours(start_date, end_date, download_folder)
            
            # downloadedSuccessful.startup()
            # downloadedSuccessful.navigate()
            downloadedSuccessful = timeclocker.downloader()
            if not downloadedSuccessful:
                return False
            
            # deprecated
            if downloadedSuccessful is None:
                return False
            
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
                try:
                    # try the new method of reading timeclock
                    times_df = new_and_imporved_group_hours_html_reader(latest_html, in_and_out_times=in_and_out_times)
                except Exception:
                    # try the old way of reading timeclock incase the top fails
                    times_df = new_output_each_clock_entry_job_and_costcode(latest_html, in_and_out_times=in_and_out_times)
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
            employee_info_csvs = [f for f in list_of_csvs if "Employee Information" in f]
            # Get the most recent file for that state
            latest_csv = max(employee_info_csvs, key=os.path.getctime)
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


def get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=True, download_folder="C:\\users\\cwilson\\downloads\\", ei=None, in_and_out_times=False):
    
    times_df = get_clock_times_html_downloaded(start_date, end_date, exclude_terminated, download_folder, in_and_out_times=in_and_out_times)
    
    
    
    
    if ei is None:
        ei = get_ei_csv_downloaded(exclude_terminated, download_folder)
    
    if isinstance(times_df, pd.DataFrame) or isinstance(times_df, pd.Series):
        return clean_up_this_gunk(times_df, ei)
    else:
        return False
    
    




























