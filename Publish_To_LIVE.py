# -*- coding: utf-8 -*-
"""
Created on Fri Mar 19 09:50:48 2021

@author: CWilson
"""


import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import gspread
from utils.google_sheets_credentials_startup import init_google_sheet
from Publish_to_live_gather_data_from_all_sources import gather_states_data
from Publish_to_live_update_history import update_history
from Grab_Fabrication_Totals import totals
import time
import datetime


states = ['TN', 'MD', 'DE']
# states = ['DE']
# states = ['MD', "TN"]

# get todays date
today = datetime.date.today()
# get the date of sunday - the start of the week
week_start = today - datetime.timedelta(days=today.weekday() + 1)
# get the date of the end of the week
week_end = week_start + datetime.timedelta(days = 6)

# get the start of the month via the today variable
month_start = datetime.date(today.year, today.month, 1)
# calculate the last day of the month by going to the next month and subtracting one day
month_end = datetime.date(today.year, today.month+1, 1) - datetime.timedelta(days=1)

# convert the dates into a string format for all my functions
today = today.strftime("%m/%d/%Y")
month_start = month_start.strftime("%m/%d/%Y")
month_end = month_end.strftime('%m/%d/%Y')

# create an empty dic to put the state data into
state_data = {}
# loops thru each state to pull all their data
for state in states:
    state_data[state] = {}

    state_data[state]['Today'] = gather_states_data(state, today, today)
    
    state_data[state]['Month'] = gather_states_data(state, month_start, month_end)



''' This snippet makes it so either all states update or none update '''
# defaults to be allowed to run
good_to_run = True
# checks if the value for each state is a list or not
# if there was an error during the 'gather_states_data()', the value becomes a NoneType
for val in state_data:
    if isinstance(state_data[val]['Today'], list):
        print(val, " is good")
    else:
        print(val, " is bad")
        # toggles the ability to run the program if there is a NoneType present
        good_to_run = False



try:
    
    if good_to_run:
        # counts the number of elements that will be updated by looping thru all the lists
        # I want to use this to change the sleep timer.
        # start count at # of states * 10
        # *2 bc date and time
        # *2 for in process (update time then update to blank)
        # *6 for year, month, week totals
        count = len(states) * 10
        for state in state_data:
            for row in state_data[state]['Today']:
                count += len(row)
            for row in state_data[state]['Month']:
                count += len(row)
        print('# of cells to be updated: ' + str(count))
        # use count to determine the sleep timer for API limits
        # setting sleep timer to 1.5 second for now b/c I don't want to figure out how it needs to be changed
        x = 1.1
        
        # the ID # for the Live sheet
        live_output_google_sheets_key = "12yFpSXyblbhueEM6e5vz_WaEXJKTsozHvkfS9gjux3w"
        # Grab the spreadsheet
        sh = init_google_sheet(live_output_google_sheets_key)
        # the individual sheet is grabbed during the loop below
    
        # loops thru each state's data
        for state in state_data:
            
            # grab the sheet name to be updating
            if state == 'TN':    
                base_sheet_name = 'CSM'
            elif state == 'MD':
                base_sheet_name = 'FED'
            elif state == 'DE':
                base_sheet_name = 'CSF'
            
            
            
            # gets the Weekly, Monthly, Yearly totals from Fab Listing
            # gets the totals before doing any of the work on the shop to speed things up
            totals_lists = totals(base_sheet_name + ' QC Form')       
            
            # open the specific shop's behind-the-scenes sheet
            worksheet = sh.worksheet(base_sheet_name)
            print("Editting " + base_sheet_name + " 'Today' behind the scenes")
            
            # Grab the date from A1
            last_date = worksheet.acell('A1').value
            # # Grab the Time from B1
            last_time = worksheet.acell('B1').value
            # Combine the date and time
            last_datetime = last_date + ' ' + last_time   
            # run the update history function which grabs stuff from the main display page
            # and moves that information to the Shop history page
            update_history(last_datetime, state)   
            # get the current datetime
            right_now = datetime.datetime.now()
            # break it into date string
            right_now_date = right_now.date().strftime("%m/%d/%Y")
            # break it into time string
            right_now_time = right_now.time().strftime("%I:%M %p")
            
            
            ''' Update the 'Today' sheet '''
            # clears the behind-the-scenes worksheet
            worksheet.clear()
            # Adds the IN PROCESS tag with a timestamp
            worksheet.update('C1', "IN PROCESS as of: " + right_now_time)       
            # pause
            time.sleep(x)
            # Also need to update cells A1 for the Date
            worksheet.update('A1', right_now_date, value_input_option='USER_ENTERED')
            # Update cell B1 with the time
            worksheet.update('B1', right_now_time, value_input_option='USER_ENTERED')
            
            # loops through the Job #, Weight, Direct Hours
            for i,row in enumerate(state_data[state]['Today']):
                # loop through the individual values in those lists
                for j, value in enumerate(row):
                    # determine amount of sleep to not trigger the API limits
                    time.sleep(x)
                    # converts a number to corresponding letter (the +97 is needed)
                    letter = chr(j + 97)
                    # the +3 moves it down to the third row of the sheet
                    number = i + 3
                    # combine the letter and number to get the cell to edit
                    cell = letter + str(number)
                    # updates the cell 
                    worksheet.update(cell, value)
            
            # loops thru each list of the week/month/year totals
            for i,level in enumerate(totals_lists):
                # the starting row should be on Row #8 in the behind-the-scenes sheet
                row = i + 8
                # loop thru each value in the list
                for j,val in enumerate(level):
                    # pause
                    time.sleep(x)
                    # the Columns should only be A & B
                    col = chr(j + 97)
                    # get the cell number
                    cell = col + str(row)
                    # update that cell
                    worksheet.update(cell, int(val))
                    
        
            # Removes the in process tag
            worksheet.update('C1', '')
            
            # ''' Update the 'Month' Sheet '''
            # print("Editting " + base_sheet_name + " 'Month' behind the scenes")
            # # load the month sheet
            # worksheet = sh.worksheet(base_sheet_name + ' Month')
            # # get the current timestamp
            # right_now = datetime.datetime.now()
            # # break it into date string
            # right_now_date = right_now.date().strftime("%m/%d/%Y")
            # # break it into time string
            # right_now_time = right_now.time().strftime("%I:%M %p")   
            # # clear the behind-the-scenes month sheet
            # worksheet.clear()
            # # Adds the IN PROCESS tag with a timestamp
            # worksheet.update('C1', "IN PROCESS as of: " + right_now_time)       
            # # pause
            # time.sleep(x)            
            # # Also need to update cells A1 for the Date
            # worksheet.update('A1', right_now_date, value_input_option='USER_ENTERED')
            # # Update cell B1 with the time
            # worksheet.update('B1', right_now_time, value_input_option='USER_ENTERED')    
            # # loops through the Job #, Weight, Direct Hours
            # for i,row in enumerate(state_data[state]['Month']):
            #     # loop through the individual values in those lists
            #     for j, value in enumerate(row):
            #         # determine amount of sleep to not trigger the API limits
            #         time.sleep(x)
            #         # converts a number to corresponding letter (the +97 is needed)
            #         letter = chr(j + 97)
            #         # the +3 moves it down to the third row of the sheet
            #         number = i + 3
            #         # combine the letter and number to get the cell to edit
            #         cell = letter + str(number)
            #         # updates the cell 
            #         worksheet.update(cell, value)        
            
            
            # # get the row number from the last row in the previous loop
            # number = number + 1
            # for col in range(0,j+1):
            #     # determine amount of sleep to not trigger the API limits
            #     time.sleep(x)                
            #     # get the column letter
            #     letter = chr(col + 97)
            #     # get the cell 
            #     cell = letter + str(number)
            #     # get the 3rd row to get the job cell
            #     job_cell = letter + str(number - 4)
            #     # get the 4th row to get the weight cell
            #     weight_cell = letter + str(number - 3)
            #     # makeup the formula string 
            #     formula = "=iferror(Vlookup(" + job_cell +"," + base_sheet_name
            #     formula += "SoldHours,2,False) / 2000 * " + weight_cell + ", 0)"                
            #     # update the cell to have the formula in it 
            #     worksheet.update(cell, formula, value_input_option='USER_ENTERED')
                
            # # "=iferror(Vlookup(E$36,CSMSoldHours,2,False) / 2000 * E$37, E$38)"
            
            # # remove the in process tag
            # worksheet.update('C1', '')
            # ''' Finished updating the 'Month' sheet '''
            
            
            
            
            
except Exception as e:
    # create a copy of the current variables in memory
    current_variables = dir().copy()
    # error log directory
    error_log = "C:\\Users\\cwilson\\Documents\\Python\\Publish to Live\\Error_logs\\"
    # gets current date to timestamp the file
    error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # create the entire file name
    file_name = error_log + "publish_to_live_error-" + error_date + ".txt"
    # open the file
    file = open(file_name, 'w')
    # write the state name to the file
    file.write(state + '\n')
    # write the error to the file
    file.write(str(e))
    file.write('\n--- Start of the Variables ---\n')
    # iterate thru each variable 
    for var_name in current_variables:
        # does not proceed to write the variable name & value if it is a module
        # if inspect.ismodule(current_variables[var_name]):
        #     continue
        # write the variable name then a linebreak
        file.write('\n' + var_name + '\n')
        
        # write the variable contents 
        # file.write(current_variables[var_name])
        # write a line break, then dashes, then another line break
        # file.write('\n--- end: ' + var_name + ' ---\n')

    # close the file
    file.close()
