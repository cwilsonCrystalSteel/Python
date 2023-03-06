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
import xlsxwriter
from openpyxl import load_workbook
code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))

base = 'c:\\users\\cwilson\\documents\\Productive_Employees_Hours_Worked_report\\'
backup = base + 'Backups\\'


states = ['TN','MD','DE']


def run_attendance_hours_report(state):
    
    now = datetime.datetime.today()
    today_stamp = datetime.datetime.strftime(now, '%Y-%m-%d %H-%M')
    
    # open up each of the excel files and figure out what the earliest possible start date could be
    earliest_start_dts = {}
    dumb_list = []

    file_name = 'week_by_week_hours_of_employees ' + state 
    file_path = base + file_name + '.xlsx'    
    starter = pd.read_excel(file_path, sheet_name='Data')
    starter = starter.set_index('Week Start')
    remove_cols = starter.columns[starter.iloc[0].isna()]
    starter = starter.drop(columns=remove_cols)
    # the summary rows are the last 5
    summary = starter.iloc[-5:]
    # get the starter df that does not contain the indexes contained in the summary
    data = starter[~starter.index.isin(summary.index)]
    # drop the na rows - the rows between the data & the summary
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
        # for state in earl1iest_start_dts.keys():
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
        start_date0 = start_dt.strftime("%m/%d/%Y")
        end_date0 = (start_dt + datetime.timedelta(days=3)).strftime("%m/%d/%Y")
        start_date1 = (start_dt + datetime.timedelta(days=4)).strftime("%m/%d/%Y")
        end_dt = start_dt + datetime.timedelta(days=6)
        end_date1 = end_dt.strftime("%m/%d/%Y")
        # get the basis dict
        basis0 = get_information_for_clock_based_email_reports(start_date0, end_date0, exclude_terminated=False)    
        basis1 = get_information_for_clock_based_email_reports(start_date1, end_date1, exclude_terminated=False) 
        # go through each state
        # for state in reversed_start_date_dict[start_dt]:

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
        ei = basis0['Employee Information']
        # set the index to the employee name
        ei = ei.set_index('Name')
        # get all employees at that state
        ei = ei[ei['Productive'].str.contains(state)]
        # Get all the hours put into one df
        hours = basis0['Direct'].append(basis0['Indirect']).append(basis1['Direct']).append(basis1['Indirect'])
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
        
        with pd.ExcelWriter(file_path) as writer:  

            data.to_excel(writer, sheet_name='Data')
        
        try:
            new_file_path = create_formatted_excel(data, file_path)
        except Exception:
            # if the formatting funciton fails - just pass the plain excel file
            new_file_path = file_path
        
        # book = load_workbook(file_path)
        # book.remove(book['Data'])
        # writer = pd.ExcelWriter(file_path, engine='openpyxl')
        # writer.book = book
        
        
        # data.to_excel(writer, sheet_name='Data')
        # writer.save()
        # writer.close()
        
        
        
    return {'filepath':new_file_path, 'weekstart':start_dt.strftime('%m/%d/%Y')}
    
    
        
        
        
    
def create_formatted_excel(xlsx_structured_df, file_path):
    
    new_file_path = file_path.split('.xlsx')
    new_file_path = new_file_path[0] + '_formatted' + '.xlsx'
    
    
    
    workbook = xlsxwriter.Workbook(new_file_path)
    worksheet = workbook.add_worksheet('Data')
    # freeze first col and first row
    worksheet.freeze_panes(1,1)
    # define row and column height for for frozen col and row
    worksheet.set_column(0,0, 17.00)
    worksheet.set_row(0, 58.00)    
    # define some traits 
    bold = workbook.add_format({'bold': True})
    wrap = workbook.add_format()
    wrap.set_text_wrap()
    
    one_decimal = workbook.add_format({'num_format': '0.0'})
    
    # red
    red = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})
    # green
    green = workbook.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100'})
    # yellow
    yellow = workbook.add_format({'bg_color': '#ffeb9c',
                                   'font_color': '#9f5700'})

    black = workbook.add_format({'bg_color': '#000000',
                                   'font_color': '#000000'})      
    
    
    
    worksheet.write(0,0, 'Week Start')
    
    col = 1
    # get the headers into the excel file
    for col_name in xlsx_structured_df.columns:
        worksheet.write(0, col, col_name, wrap)
        col += 1
    
    
    
    row = 1
    # get the index into the excel file
    for index_name in xlsx_structured_df.index:
        
        if isinstance(index_name, datetime.datetime):
                index_name = index_name.strftime('%Y-%m-%d')
                
        worksheet.write(row, 0, index_name, bold)
        row += 1
        

    
    # workbook.close()
    
    row, col = 1, 1
    
    
    # get all of the values into the file 
    for y in range(0, xlsx_structured_df.shape[0]):
        for x in range(0, xlsx_structured_df.shape[1]):
            
            value = xlsx_structured_df.iloc[y, x]
            if isinstance(value, datetime.datetime):
                value = value.strftime('%Y-%m-%d')
                worksheet.write(row + y, col + x, value)
            if pd.isna(value):
                value = ''
                worksheet.write(row + y, col + x, value)
            else:
            
                worksheet.write(row + y, col + x, value, one_decimal)
    
    # workbook.close()
    
    col_letter_start = 'F'
    if x + 1 <= 26:
        col_letter_end = chr(ord('A') + x + 1)
    else:
        #### This was givving me an invalivd column letter when x=103 -> C[
        # col_letter_end = chr(ord('A') +  int(x / 26) - 1) + chr(ord('A') + int(x % 26) + 1)
        #### switching to this
        x += 1
        col_letter_end = chr(ord('A') +  int(x / 26) - 1) + chr(ord('A') + int(x % 26))
        
        
    row_number_start_body = '2'
    row_number_end_body = str(y - 6)
    
    row_number_start_summary = str(int(row_number_end_body) + 4)
    row_number_end_summary = str(y + 2)
    
    cell_start_body = col_letter_start + row_number_start_body
    cell_end_body = col_letter_end + row_number_end_body
    cell_range_body = cell_start_body + ':' + cell_end_body
    
    cell_start_summary = col_letter_start + row_number_start_summary
    cell_end_summary = col_letter_end + row_number_end_summary
    cell_range_summary = cell_start_summary + ':' + cell_end_summary
    
  
    
    worksheet.conditional_format(cell_range_body, {'type': 'cell',
                                      'criteria': '=',
                                      'value': 0,
                                      'format': black})    
    
    worksheet.conditional_format(cell_range_body, {'type': 'cell',
                                      'criteria': '>=',
                                      'value': 48,
                                      'format': green})
    
    worksheet.conditional_format(cell_range_body, {'type': 'cell',
                                      'criteria': '<=',
                                      'value': 40,
                                      'format': red})
    
    worksheet.conditional_format(cell_range_body, {'type': 'cell',
                                      'criteria': '<',
                                      'value': 48,
                                      'format': yellow}) 
    
    worksheet.conditional_format(cell_range_summary, {'type': 'cell',
                                      'criteria': '=',
                                      'value': 0,
                                      'format': black})       
    
    worksheet.conditional_format(cell_range_summary, {'type': 'cell',
                                      'criteria': '>=',
                                      'value': 48,
                                      'format': green})
    
    worksheet.conditional_format(cell_range_summary, {'type': 'cell',
                                      'criteria': '<=',
                                      'value': 40,
                                      'format': red})    
    
    worksheet.conditional_format(cell_range_summary, {'type': 'cell',
                                      'criteria': '<',
                                      'value': 48,
                                      'format': yellow})
    
    
    
    
    
    
    workbook.close()
    
    
    return new_file_path
    


    






