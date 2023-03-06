# -*- coding: utf-8 -*-
"""
Created on Mon Feb  7 17:08:50 2022

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










for state in states:
    
    file_name = 'week_by_week_hours_of_employees ' + state 
    file_path = base + file_name + '.xlsx'
    backup_file_path = backup + file_name + ' ' +today_stamp + '.xlsx'
    
    copyfile(file_path, backup_file_path)
    print('\nCopy of file made before new week added: ')
    print(backup_file_path, end='\n\n')
    
    
    starter = pd.read_excel(file_path, sheet_name='Data')
    starter = starter.set_index('Week Start')
    remove_cols = starter.columns[starter.iloc[0].isna()]
    starter = starter.drop(columns=remove_cols)
    summary = starter.iloc[-5:]
    data = starter[~starter.index.isin(summary.index)]
    data = data.dropna()
    
    
    # start_date = "05/02/2021"
    # start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
    start_dt = max(data.index) + datetime.timedelta(days=7)
    start_date = start_dt.strftime("%m/%d/%Y")
    end_dt = start_dt + datetime.timedelta(days=6)
    end_date = end_dt.strftime("%m/%d/%Y")
    
    basis = get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=False)
    ei = basis['Employee Information']
    ei = ei.set_index('Name')
    # 
    prod_code = state + ' PRODUCTIVE'
    # ei = ei[ei['Productive'] == prod_code]
    ei = ei[ei['Productive'].str.contains(state)]
    hours = basis['Direct'].append(basis['Indirect'])
    hours = hours[~hours['Job Code'].isin(code_changes['Delete Job Codes'])]
    hours = hours.groupby('Name').sum()
    hours = hours.join(ei['Productive'])
    hours = hours[~hours['Productive'].isna()]
    hours = hours[['Hours']].transpose()
    hours['Week Start'] = start_dt
    hours = hours.set_index('Week Start')
    
    
    hours_worked = hours.sum(axis=1)[0]
    number_worked = hours[hours>0].count(axis=1)[0]
    goal_worked = 48 * number_worked
    missing_hours = goal_worked - hours_worked
    if missing_hours < 0:
        missing_hours = 0
    
    hours_plus = hours.copy()
    
    hours_plus['Hours Worked'] = hours_worked
    hours_plus['Num. Worked'] = number_worked
    hours_plus['48 x Num. Worked'] = goal_worked
    hours_plus['Missing Hours'] = missing_hours
    
    data = data.append(hours_plus)
    data = data.fillna(0)
    
    summary = summary.copy()
    summary.loc['Average'] = data.mean()
    summary.loc['Average (if worked)'] = data[data > 0].mean()
    summary.loc['12-Week Average'] = data.iloc[-12:][data > 0].mean()
    summary.loc['8-Week Average'] = data.iloc[-8:][data > 0].mean()
    summary.loc['4-Week Average'] = data.iloc[-4:][data > 0].mean()
    summary = summary.fillna(0)
    
    
    
    
    for i in range(0,3):
        # data.loc[data.shape[0]] = [None] * data.shape[1]
        data = data.append(pd.Series(name='', dtype=float))
    
    data = data.append(summary)
    
    columns_start = ['Hours Worked', 'Num. Worked', '48 x Num. Worked', 'Missing Hours']
    
    # sort the hours by most to least worked
    hours = hours.squeeze().sort_values(ascending=False).to_frame().transpose()
    
    columns_rest = list(hours.columns)
    
    columns_missing = [i for i in list(data.columns) if i not in list(hours_plus.columns)]
    
    columns_order = columns_start + columns_rest + columns_missing
    
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
