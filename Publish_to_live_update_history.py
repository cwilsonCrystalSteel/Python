# -*- coding: utf-8 -*-
"""
Created on Thu Mar 11 09:22:30 2021

@author: CWilson
"""

import datetime
import time


def update_history(last_datetime, state):
    
    from utils.google_sheets_credentials_startup import init_google_sheet

    live_output_google_sheets_key = "12yFpSXyblbhueEM6e5vz_WaEXJKTsozHvkfS9gjux3w"
    # Grab the spreadsheet
    sh = init_google_sheet(live_output_google_sheets_key)
    # Grab from the Combined Sheet
    worksheet = sh.worksheet('Combined')    
    
    
    
    # grabs the Combined sheet as a list of lists
    combined = worksheet.get_all_values()
    # checks each row to see if it is blank or not
    for i,row in enumerate(combined):
        if len(row[0]) != 0:
            # The index of the first non-blank row is grabbed 
            starter = i 
            break
    
    
    

    # Chooses the history sheet for updating
    # Chooses the first row to start pulling data from on the Combined sheet
    if state == 'TN':
        sheet_name = 'CSM History'
        start_row = starter + 4
    elif state == 'MD':
        sheet_name = 'FED History'
        start_row = starter + 22
    elif state == 'DE':
        sheet_name = 'CSF History'
        start_row = starter + 13
        
        
    
    
    
    
    # Initiliaze the list with the datetime
    vals = [last_datetime]
    # grab the DL Efficiency, Total Tons, Total Direct, Direct Hrs/Ton, Total Earned
    for i in range(0,5):
        val = worksheet.acell('B' + str(start_row + i)).value
        vals.append(val)
    
    # Switches to the History sheet
    worksheet = sh.worksheet(sheet_name)
    
    
    # Find the current last row
    values = worksheet.get_all_values()
    # Grabs the number of the new row to post data to
    row_num = len(values)
    # if the history sheet is freshly wiped by archive_history_as_csv.py
    # this moves the starting row down one more so it does not wipe out
    # the headers when it puts the first row of data in place
    if row_num == 1:
        row_num = 2
    
    # pastes the new data into the sheet
    for i, val in enumerate(vals):
        time.sleep(1)
        cell = chr(i+97) + str(row_num)
        worksheet.update(cell, val, value_input_option='USER_ENTERED')
        
    
    
    '''
    attempt to get the last row of the history sheet to grab the current values
    from the combined sheet and display them as the last row of history
    this will then let you view the current times data on the charts
    '''
    
    # row with the "Last update %date %time" is 2 rows before start_row
    combined_datetime_row = start_row - 3
    # grabs the cell with the date in it
    combined_date_cell = "Combined!F" + str(combined_datetime_row)
    # grabs the cell with the time in it
    combined_time_cell = "Combined!G" + str(combined_datetime_row)
    # Create the string that is the formula to combine the date and time 
    date_formula  = "=VALUE(TEXT(" + combined_date_cell + ", \"mm/dd/yyy\") & \" \" & TEXT(" + combined_time_cell + ", \"H:MM AM/PM\"))"
    # the direct labor formula is just a cell reference
    dl_formula = "=Combined!B" + str(start_row)
    # The tons formula is just a cell reference
    tons_formula = "=Combined!B" + str(start_row + 1)
    # The direct hours formula is just a cell reference
    direct_formula = "=Combined!B" + str(start_row + 2)
    # the earned hours formula is just a cell reference
    earned_formula = "=Combined!B" + str(start_row + 3)
    # the eva hours formula is just a cell reference
    eva_formula = "=Combined!B" + str(start_row + 4)
    
    # combine all the formulas into a string to loop thru
    last_row_formulas = [date_formula, dl_formula, tons_formula, direct_formula, earned_formula, eva_formula]
    
    # iterate thru the formulas
    for i,formula in enumerate(last_row_formulas):
        # sleep for API limits
        time.sleep(1)
        # get the column letter with chr(i+97)
        # put the last row as the row with the formulas -> str(row_num + 1)
        cell = chr(i+97) + str(row_num + 1)
        # update the cell
        worksheet.update(cell, formula, value_input_option='USER_ENTERED')
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    