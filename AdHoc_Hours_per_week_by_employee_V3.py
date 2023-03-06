# -*- coding: utf-8 -*-
"""
Created on Tue Aug 10 15:04:26 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
from shutil import copyfile
from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
import pandas as pd
import datetime
import json
from openpyxl import load_workbook
code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))

base = 'c:\\users\\cwilson\\documents\\Productive_Employees_Hours_Worked_report\\'
backup = base + 'Backups\\'




x = input("What states should be ran? 'TN', 'DE', 'MD', or 'ALL'? ")
for i in range(0,6):
    
    x = 'ALL'
        
    states = ['TN','MD','DE']
    
    if x in states:
        states = [x]
    elif x == 'ALL':
        states = states
    else:
        print('Not a valid selection, please close & try again')
        quit()
    
    
    
    
    now = datetime.datetime.today()
    today_stamp = datetime.datetime.strftime(now, '%Y-%m-%d %H-%M')
    
    # open up each of the excel files and figure out what the earliest possible start date could be
    earliest_start_dts = {}
    dumb_list = []
    for state in states:
        file_name = 'week_by_week_hours_of_employees ' + state 
        file_path = base + file_name + '.xlsx'    
        starter = pd.read_excel(file_path, sheet_name='Data')
        starter = starter.set_index('Week Start')
        remove_cols = starter.columns[starter.iloc[0].isna()]
        starter = starter.drop(columns=remove_cols)
        summary = starter.iloc[-5:]
        data = starter[~starter.index.isin(summary.index)]
        data = data.dropna()
        start_dt = max(data.index) + datetime.timedelta(days=7)
        earliest_start_dts[state] = start_dt
        dumb_list.append(start_dt)
    
    
    # init a new dict that will have keys being the possible start dates
    reversed_start_date_dict = {}
    # iterate thru each possible start date found from the current excel files
    for start_dt in list(set(dumb_list)):
        # dont add the start date if it greater than today's date
        if start_dt > now:
            continue
        
        # loop thru each state in the earliest_start_dts dict
        for state in earliest_start_dts.keys():
            # check if the list of states needs to be initalized to that start_dt key
            if earliest_start_dts[state] == start_dt and start_dt not in reversed_start_date_dict:
                reversed_start_date_dict[start_dt] = [state]
            
            # append to the list bc the list is already made for that start dt key
            elif earliest_start_dts[state] == start_dt and start_dt in reversed_start_date_dict:
                reversed_start_date_dict[start_dt].append(state)
                
    
    if len(list(reversed_start_date_dict.keys())) == 0:
        print('There are no start dates that are befoer the current date - ENDING PROGRAM')
    
    # go through each start date
    for start_dt in reversed_start_date_dict.keys():
        
        # get the start and end dates for the basis dict
        start_date = start_dt.strftime("%m/%d/%Y")
        end_dt = start_dt + datetime.timedelta(days=6)
        end_date = end_dt.strftime("%m/%d/%Y")
        # get the basis dict
        basis = get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=False)    
        
        # go through each state
        for state in reversed_start_date_dict[start_dt]:
    
                # get the excel file name & backup file name
                file_name = 'week_by_week_hours_of_employees ' + state 
                file_path = base + file_name + '.xlsx'
                backup_file_path = backup + file_name + ' ' +today_stamp + '.xlsx'
                
                # copy the current file & move to the backup destination
                copyfile(file_path, backup_file_path)
                print('\nCopy of file made before new week added: ')
                print(backup_file_path, end='\n\n')
                
                # read the file
                starter = pd.read_excel(file_path, sheet_name='Data')
                # change the index to the week start date
                starter = starter.set_index('Week Start')
                # get the columns that need to be removed
                remove_cols = starter.columns[starter.iloc[0].isna()]
                # remove those columns
                starter = starter.drop(columns=remove_cols)
                # grab the summary portion of the file as last 5 rows
                summary = starter.iloc[-5:]
                # grab the data portion of the excel file as not the summary part
                data = starter[~starter.index.isin(summary.index)]
                # drop any na
                data = data.dropna()
                
                # get the employee information df
                ei = basis['Employee Information']
                # set the index to the employee name
                ei = ei.set_index('Name')
                # get all employees at that state
                ei = ei[ei['Productive'].str.contains(state)]
                # Get all the hours put into one df
                hours = basis['Direct'].append(basis['Indirect'])
                # get rid of the delete job codes 
                hours = hours[~hours['Job Code'].isin(code_changes['Delete Job Codes'])]
                # group by the index & sum
                hours = hours.groupby('Name').sum()
                # Join the productive column to hours
                hours = hours.join(ei['Productive'])
                # get rid of anybody where the productive code is NA -> basically only getting that states employee hours
                hours = hours[~hours['Productive'].isna()]
                # transpose the df for the excel formatting
                hours = hours[['Hours']].transpose()
                # add column for week start
                hours['Week Start'] = start_dt
                # set the index as week start
                hours = hours.set_index('Week Start')
                
                # sum the total hours worked by all employees
                hours_worked = hours.sum(axis=1)[0]
                # count the number of employees that have worked
                number_worked = hours[hours>0].count(axis=1)[0]
                # calculate the goal number of hours 
                goal_worked = 48 * number_worked
                # subtract to get the missing # of hours
                missing_hours = goal_worked - hours_worked
                # if they work greater than the goal number then set it to zero
                if missing_hours < 0:
                    missing_hours = 0
                
                
                # copy the hours df
                hours_plus = hours.copy()
                # set the values for the columns
                hours_plus['Hours Worked'] = hours_worked
                hours_plus['Num. Worked'] = number_worked
                hours_plus['48 x Num. Worked'] = goal_worked
                hours_plus['Missing Hours'] = missing_hours
                # append the row to the data df
                data = data.append(hours_plus)
                # fill any missing values with zero
                data = data.fillna(0)
                # these are the calculated rows that average numbers for each employee
                summary = summary.copy()
                # get the total average
                summary.loc['Average'] = data.mean()
                # get the average of present weeks
                summary.loc['Average (if worked)'] = data[data > 0].mean()
                # get the 12 week average when they have worked
                summary.loc['12-Week Average'] = data.iloc[-12:][data > 0].mean()
                # get the 8 week average when they worked
                summary.loc['8-Week Average'] = data.iloc[-8:][data > 0].mean()
                # get the 4 week average when they worked
                summary.loc['4-Week Average'] = data.iloc[-4:][data > 0].mean()
                # fill any missing with zero
                summary = summary.fillna(0)
                
                
                
                # append 3 blank rows to the data ddf -> formatting for excel
                for i in range(0,3):
                    # data.loc[data.shape[0]] = [None] * data.shape[1]
                    data = data.append(pd.Series(name='', dtype=float))
                
                # now append the summary df to data
                data = data.append(summary)
                # get the columns we want to show up first / on the left
                columns_start = ['Hours Worked', 'Num. Worked', '48 x Num. Worked', 'Missing Hours']
                # sort the employees by most to least worked for the most current week
                hours = hours.squeeze().sort_values(ascending=False).to_frame().transpose()
                # get the rest of the columns
                columns_rest = list(hours.columns)
                # get any not worked employees so they dont get axed from the excel file
                columns_missing = [i for i in list(data.columns) if i not in list(hours_plus.columns)]
                # combine the columns to one list
                columns_order = columns_start + columns_rest + columns_missing
                # now order the data df columns as desired
                data = data[columns_order]
                
                book = load_workbook(file_path)
                book.remove(book['Data'])
                writer = pd.ExcelWriter(file_path, engine='openpyxl')
                writer.book = book
                
                
                data.to_excel(writer, sheet_name='Data')
                writer.save()
                writer.close()
    
        
        
            
            
            
            
        
        
        
        
    
    


while True:
    x = input('\nType "QUIT" to exit the prompt.\t')
    if x == 'QUIT':
        quit()

    



