# -*- coding: utf-8 -*-
"""
Created on Sun Aug  4 09:07:54 2024

@author: CWilson
"""

# -*- coding: utf-8 -*-
"""
Created on Thu May  6 11:09:44 2021

@author: CWilson
"""

import pandas as pd
import glob
import os
from pathlib import Path
import datetime
from Read_Group_hours_HTML import new_output_each_clock_entry_job_and_costcode, new_and_imporved_group_hours_html_reader
import json
from TimeClock.TimeClockNavigation import TimeClockBase
from TimeClock.pullEmployeeInformationFromSQL import return_sql_ei

default_download_folder = Path.home() / 'downloads' / 'EmployeeInformation'

def clean_up_this_gunk(times_df, ei):
    
    
    # rename the columns b/c it makes it easier to work with
    ei = ei.rename(columns={'<NUMBER>':'ID',
                     '<FIRSTNAME>':'First',
                     '<LASTNAME>':'Last',
                     '<LOCATION>':'Location',
                     '<CLASS>':'Shift',
                     '<SCHEDULEGROUP>':'Productive',
                     '<DEPARTMENT>':'Department'})

    ei = ei.rename(columns={'employeeidnumber':'ID',
                     'firstname':'First',
                     'lastname':'Last',
                     'location':'Location',
                     'classshift':'Shift',
                     'schedulegroup':'Productive',
                     'department':'Department'})    
        
        
        
    print('')
    # get active employees only for this 
    ei = ei[ei['terminated'] == False]
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
    


def get_clock_times_html_downloaded(start_date, end_date, exclude_terminated=True, download_folder="C:\\users\\cwilson\\downloads\\GroupHours\\", in_and_out_times=False):

    today = datetime.datetime.today().date()
    
    count = 0
    while count < 3:
        try:
            count += 1
            
            x = TimeClockBase(download_folder, offscreen=True)     
            x.startupBrowser()
            x.tryLogin()
            x.openTabularMenu()
            

            x.searchFromTabularMenu('Group Hours')
            x.clickTabularMenuSearchResults('Hours > Group Hours')
            try:
                x.groupHoursFinale(start_date)
                filepath = x.retrieveDownloadedFile(15, '*.html', 'Hours')
                print(filepath)
            except Exception as e:
                print(f'Could not complete download of {start_date} because {e}')
                
                
                ''' CREATE FUNCTION TO GENERATE EMPTY TIMES_DF #?? '''
                
                
                
                # return False
                
            try:
                # try the new method of reading timeclock
                times_df = new_and_imporved_group_hours_html_reader(filepath, in_and_out_times=in_and_out_times)
            except Exception:
                # try the old way of reading timeclock incase the top fails
                times_df = new_output_each_clock_entry_job_and_costcode(filepath, in_and_out_times=in_and_out_times)
            
            
            # delete the html file
            os.remove(filepath)   
            
            x.kill()
            
            return times_df
            
            
           
        except Exception as e:
            print('\n\nDownloading group hours failed: ')
            print(e)
            print('\n\n')
            pass
        
    if 'x' in locals():
        try:
            x.kill()
            print('retrieving times_df failed, but we are gonna kill the browser now')

        except:
            pass
        
        
    return False


def get_ei_csv_downloaded(exclude_terminated, download_folder=default_download_folder):
    ''' try to get from sql first! '''
    try:
        ei = return_sql_ei()
        return ei
    except:
        print('Could not retrieve ei from SQL')
    
    
    ''' cant get it from sql? use timeclock '''
    today = datetime.datetime.today().date()
    # latest_csv = "c://users//cwilson//downloads//f_this.csv"
    count = 0
    while count < 6:
        try:
            count += 1
            
            x = TimeClockBase(download_folder, headless=True)     
            x.startupBrowser()
            x.tryLogin()
            x.openTabularMenu()
            x.searchFromTabularMenu('export')
            x.clickTabularMenuSearchResults('Tools > Export')
            try:
                x.employeeLocationFinale()
                filepath = x.retrieveDownloadedFile(20, '*.csv', 'Employee Information')
                print(filepath)
            except Exception as e:
                print(f'Could not complete download because {e}')
                
                
                
            ei = pd.read_csv(filepath)
            # Delete the Employee Information CSV file from the downloads folder
            os.remove(filepath)  
            # be done with the browser 
            x.kill()
            
            return ei
            
           
        except Exception as e:
            print('Downloading employee information failed: ')
            print(e)
            pass
    
    
    if 'x' in locals():
        try:
            x.kill()
            print('retrieving ei failed, but we are gonna kill the browser now')
        except:
            pass    
    
    return False


def get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=True, download_folder="C:\\users\\cwilson\\downloads\\", ei=None, in_and_out_times=False):
    
    times_df = get_clock_times_html_downloaded(start_date, end_date, exclude_terminated, in_and_out_times=in_and_out_times)
    
    
    
    
    if ei is None:
        ei = get_ei_csv_downloaded(exclude_terminated, download_folder)
    
    if isinstance(times_df, pd.DataFrame) or isinstance(times_df, pd.Series):
        return clean_up_this_gunk(times_df, ei)
    else:
        return False
    
    




























