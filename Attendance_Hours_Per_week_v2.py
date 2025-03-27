# -*- coding: utf-8 -*-
"""
Created on Tue Aug 10 15:04:26 2021

@author: CWilson
"""

from shutil import copyfile
from TimeClock.Gather_data_for_timeclock_based_email_reports_SQL import get_information_for_clock_based_email_reports
from TimeClock.pullGroupHoursFromSQL import get_timesdf_from_vClocktimes, get_date_range_timesdf_controller
from TimeClock.functions_TimeclockForSpeedoDashboard import return_information_on_clock_data

import pandas as pd
import datetime
import json
import xlsxwriter
from openpyxl import load_workbook
import os
from pathlib import Path

code_changes_file = Path(os.getcwd()) / 'job_and_cost_code_changes.json'

code_changes = json.load(open(code_changes_file))


base = Path.home() / 'documents' / 'Attendance_Hours_Worked'
backup = base / 'Backups'
if not os.path.exists(base):
    os.makedirs(base)
if not os.path.exists(backup):
    os.makedirs(backup)


states = ['TN','MD','DE']


def run_attendance_hours_report(state):
    
    now = datetime.datetime.today()
    
    # Find the most recent Sunday
    last_sunday = now - datetime.timedelta(days=now.weekday() + 1)
    
    if (now-last_sunday).days < 7:
        last_sunday = last_sunday - datetime.timedelta(days = 7)
        
    last_sunday = last_sunday.date()
    
    # Find the Sunday before that - to update it with remediated values!
    previous_sunday = last_sunday - datetime.timedelta(days=7)
    
    file_name = 'Weekly Hours Worked By Employees - ' + state + '.xlsx'
    file_path = base / file_name    
    
    # if we dont even have a file, lets start here
    if not os.path.exists(file_path):
        start_dt = datetime.date(2022,6,19) # earlist available in db as of 2025-03-25
        # start_dt = datetime.date(2023,12,31) # happens to be a sunday
        
        # start_dt = datetime.date(2025,2,23)
        # list of weeks we will run
        weeks_to_run = [start_dt]
        # add 7 days 
        missing_week_start_dt = start_dt + datetime.timedelta(days=7)
        # loop until we hit the max
        while missing_week_start_dt <= last_sunday:
            # add to the list
            weeks_to_run += [missing_week_start_dt]
            # increase by a week
            missing_week_start_dt += datetime.timedelta(days = 7)
            
        
        
        
    # this is so that we can do some catchup
    else:
        
        starter = pd.read_excel(file_path, sheet_name='Data', index_col='Week Start')
        # get all the datetime values from the index
        weeks_ran = [i.date() for i in starter.index if isinstance(i,datetime.datetime)]
        
        weeks_to_run = []
        # if we see the previous_sunday in there, we know it was probably run before being finalized
        if previous_sunday in weeks_ran:
            weeks_to_run += [previous_sunday]
            print(f"We will remediate the week of {previous_sunday.strftime('%Y-%m-%d')}")
            
        # start off with where the file was left off at 
        missing_week_start_dt = max(weeks_ran) + datetime.timedelta(days=7)
        while missing_week_start_dt <= last_sunday:
            # add to the list
            weeks_to_run += [missing_week_start_dt]
            # increase by a week
            missing_week_start_dt += datetime.timedelta(days = 7)
        
        
    print(f"Found the following dates to run: {', '.join([i.strftime('%Y-%m-%d') for i in weeks_to_run])}")
    
    # go through each start date
    for start_dt in weeks_to_run:
        
        print(f"Running {state} for {start_dt.strftime('%Y-%m-%d')}")
        
        start_date = start_dt.strftime("%m/%d/%Y")
        end_dt = start_dt + datetime.timedelta(days=6)
        end_date = end_dt.strftime("%m/%d/%Y")
        
        times_df = get_timesdf_from_vClocktimes(start_date, end_date)
        basis = return_information_on_clock_data(times_df)
        
      

        
        # get the employee information df
        # ei = basis0['Employee Information']
        ei = basis['Employee Information']
        # set the index to the employee name
        ei = ei.set_index('Name')
        # get all employees at that state
        ei = ei[ei['Productive'].str.contains(state)]
        # Get all the hours put into one df
        # hours = basis0['Direct'].append(basis0['Indirect']).append(basis1['Direct']).append(basis1['Indirect'])
        hours = pd.concat([basis['Direct'], basis['Indirect']])
        
        hours['Hours'] = hours['Hours'].astype(float)
        
        if 'Time In' in hours.columns:
            hours = hours.drop(columns=['Time In'])
        if 'Time Out' in hours.columns:
            hours = hours.drop(columns=['Time Out'])
        
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
        hours_worked = hours.sum(axis=1).iloc[0]
        # count the number of employees that have worked
        number_worked = hours[hours>0].count(axis=1).iloc[0]
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
        
        # create the file name
        file_name = 'Weekly Hours Worked By Employees - ' + state + '.xlsx'
        # make it a filepath
        file_path = base / file_name
        today_stamp = datetime.datetime.strftime(now, '%Y%m%d%H%M%S')
        # create the backup files name
        backup_file_path = backup / (file_name + ' ' + today_stamp + '.xlsx')
        
        
        
        # we do not have a file yet, so lets create it
        if not os.path.exists(file_path):
            # copy the hours plus df
            data = hours_plus.copy()
            # replace any missing values with 0
            data = data.fillna(0)
            # make sure it is ordered correctly - yet this is not a problem for when the file doesnt exist yet
            data = data.sort_index()
            # these are the calculated rows that average numbers for each employee
            summary = data.copy()
            # we dont want the individual week's data in the summary table
            summary = summary.drop(data.index)
           
            
        
        # we already have a file, so we need to read it and update it
        else:
        # copy the current file & move to the backup destination
            copyfile(file_path, backup_file_path)
            # print('\nCopy of file made before new week added: ')
            # print(backup_file_path, end='\n\n')
        
            # read the file
            starter = pd.read_excel(file_path, sheet_name='Data', index_col='Week Start')
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
            # try and force index to be datetime.date
            data.index = pd.to_datetime(data.index).date
            # only want to try and drop out the previous sunday to Recalculate it if its in the file already & we want to run it with weeks_to_run
            # we could also do: if start_dt == previous_sunday
            if previous_sunday in data.index and previous_sunday in weeks_to_run:
                # if we are doing a remediation of previous_sunday
                data = data.drop(index=previous_sunday)
          
            # append the row to the data df
            data = pd.concat([data, hours_plus])
            # sort it after potentially doing an out of order for the remediation
            data = data.sort_index()
            
            # these are the calculated rows that average numbers for each employee
            summary = summary.copy()
            # figure out which names are in data (could be new for this week, from df:hours) but not in summary
            summary_missing_cols = list(set(data.columns) - set(summary.columns))
            
            summary[summary_missing_cols] = float(0)
            
        
        ''' this is all the same if the file is new or the file alraedy has data '''
            
        # fill any missing values with zero
        data = data.fillna(0)
        # get the total average
        summary.loc['Average'] = data.mean()
        # get the average of present weeks
        summary.loc['Average (if worked)'] = data[data > 0].mean()
        # get the 12 week average when they have worked
        # summary.loc['12-Week Average'] = data.iloc[-12:][data > 0].mean()
        summary.loc['12-Week Average'] = data.iloc[-12:].where(data.iloc[-12:] > 0).mean()
        # get the 8 week average when they worked
        # summary.loc['8-Week Average'] = data.iloc[-8:][data > 0].mean()
        summary.loc['8-Week Average'] = data.iloc[-8:].where(data.iloc[-8:] > 0).mean()
        # get the 4 week average when they worked
        # summary.loc['4-Week Average'] = data.iloc[-4:][data > 0].mean()
        summary.loc['4-Week Average'] = data.iloc[-4:].where(data.iloc[-4:] > 0).mean()

        # fill any missing with zero
        summary = summary.fillna(0)
            
            
        # Create an empty DataFrame with the same columns
        empty_rows = pd.DataFrame([[None] * data.shape[1]] * 3, columns=data.columns, dtype=float)
        
        # Append empty rows to the original DataFrame
        output = pd.concat([data, empty_rows, summary], ignore_index=False)
            
        # this is for excel formatting
        # Create an empty DataFrame with the same columns
        empty_rows = pd.DataFrame([[None] * data.shape[1]] * 3, columns=data.columns, dtype=float, index=[""]*3)
        # Append empty rows and the summary to the data df
        output = pd.concat([data, empty_rows, summary], ignore_index=False)
            
        # get the columns we want to show up first / on the left
        columns_start = ['Hours Worked', 'Num. Worked', '48 x Num. Worked', 'Missing Hours']
        # sort the employees by most to least worked for the most current week
        hours = hours.squeeze().sort_values(ascending=False).to_frame().transpose()
        # get the rest of the columns
        columns_rest = list(hours.columns)
        # get any not worked employees so they dont get axed from the excel file
        columns_missing = [i for i in list(output.columns) if i not in list(hours_plus.columns)]
        # combine the columns to one list
        columns_order = columns_start + columns_rest + columns_missing
        # now order the output df columns as desired
        output = output[columns_order]
        
        output.index.name = 'Week Start'
            
            
            
        # output = output.reset_index(drop=False)    
        with pd.ExcelWriter(file_path) as writer:  
            output.to_excel(writer, sheet_name='Data', index=True)
        
        try:
            new_file_path = create_formatted_excel(output, file_path)
        except Exception:
            # if the formatting funciton fails - just pass the plain excel file
            new_file_path = file_path
    
        
        
        
    return {'filepath':new_file_path, 'weekstart':start_dt.strftime('%m/%d/%Y')}
    
    
        
        
        
    
def create_formatted_excel(xlsx_structured_df, file_path):
    
    new_file_name = os.path.basename(file_path).split('.xlsx')[0] + ' - formatted.xlsx'
    new_file_path = Path(os.path.dirname(file_path)) / new_file_name
    
    
    
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
        
        if isinstance(index_name, datetime.datetime) or isinstance(index_name, datetime.date):
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
    


    






