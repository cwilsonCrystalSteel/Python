# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 14:02:51 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import datetime
import gspread
from utils.google_sheets_credentials_startup import init_google_sheet
from Update_History_Google_Sheet import update_history
import time
import timeit



states = ['TN', 'MD', 'DE']
# Attempt to prevent constant data loss of MD and DE
# import random
# random.shuffle(states)
# states = ['DE', 'MD']


def gather_states_data(state):


# for state in states:
    
    # wrapped the whole thing in a try-except in an attempt to limit failed attempts
    # Also logs the errors 
    try:
        start_time = timeit.default_timer()
        
        print('Started ' + state)

        # get todays date in string %m/%d/%y format to run timeclock on 
        today = datetime.date.today().strftime("%m/%d/%Y")
        # Grabs the latest data from the Daily Fab Listing
        # Only takes the last 100 entries and then takes only todays work
        from Combine_HTML_and_SHEET import return_combined_hours_weights
        # Returns the direct labor df, date of update, time of update
        combined_data = return_combined_hours_weights(state, start_date=today, end_date=today)
        # The first item in the list is todays direct hours
        df = combined_data[0]
        
        # First row of data - job #'s
        row1 = df.columns.tolist()
        # Second row of data - weights on that job
        row2 = df.loc['Weight'].tolist()
        # Third row of data - hours on that job
        row3 = df.loc['Hours'].tolist()
        
        print('Finished: ' + state)
        elapsed = timeit.default_timer() - start_time
        print(elapsed)
        
        
        
        return [row1, row2, row3]
    
    except Exception as e:
        # error log directory
        error_log = "C:\\Users\\cwilson\\Documents\\Python\\Error_logs\\"
        # gets current date to timestamp the file
        error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")
        # create the entire file name
        file_name = error_log + "error " + error_date + ".txt"
        # open the file
        file = open(file_name, 'w')
        # write the state name to the file
        file.write(state + '\n')
        # write the error to the file
        file.write(str(e))
        # close the file
        file.close()
        








# create an empty dic tot put the state data into
b = {}
# loops thru each state to pull all their data
for state in ['TN','MD','DE']:
    b[state] = gather_states_data(state)



''' This snippet makes it so either all states update or none update '''
# defaults to be allowed to run
good_to_run = True
# checks if the value for each state is a list or not
# if there was an error during the 'gather_states_data()', the value becomes a NoneType
for val in b:
    if isinstance(b[val], list):
        print(val, " is good")
    else:
        print(val, " is bad")
        # toggles the ability to run the program if there is a NoneType present
        good_to_run = False



if good_to_run:
    # counts the number of elements that will be updated by looping thru all the lists
    # I want to use this to change the sleep timer.
    # start count at 12 b/c the dates and time are updated and the inprocess cell updated 2x
    count = 12
    for state in b:
        for row in b[state]:
            count += len(row)
    print('# of cells to be updated: ' + str(count))
    # use count to determine the sleep timer for API limits
    # setting sleep timer to 1.5 second for now b/c I don't want to figure out how it needs to be changed
    x = 1.5
    
    # the ID # for the Live sheet
    live_output_google_sheets_key = "12yFpSXyblbhueEM6e5vz_WaEXJKTsozHvkfS9gjux3w"
    # Grab the spreadsheet
    sh = init_google_sheet(live_output_google_sheets_key)
    # the individual sheet is grabbed during the loop below

    # loops thru each state's data
    for state in b:
        
        # grab the sheet name to be updating
        if state == 'TN':    
            sheet_name = 'CSM'
        elif state == 'MD':
            sheet_name = 'FED'
        elif state == 'DE':
            sheet_name = 'CSF'
        
        # open the sheet
        worksheet = sh.worksheet(sheet_name)
        print("Editting " + sheet_name + " behind the scenes")
        
        
        # Grab the date from A1
        last_date = worksheet.acell('A1').value
        # # Grab the Time from B1
        last_time = worksheet.acell('B1').value
        # Combine the date and time
        last_datetime = last_date + ' ' + last_time   
        # run the update history
        update_history(last_datetime, state)   
        # get the current datetime
        right_now = datetime.datetime.now()
        # break it into date string
        right_now_date = right_now.date().strftime("%m/%d/%Y")
        # break it into time string
        right_now_time = right_now.time().strftime("%I:%M %p")
        
        # clears the behind-the-scenes worksheet
        worksheet.clear()
        # Adds the IN PROCESS tag with a timestamp
        worksheet.update('C1', "IN PROCESS as of: " + right_now_time)        
        time.sleep(x)
        
        # Also need to update cells A1 & B1 - Date and Time
        worksheet.update('A1', right_now_date, value_input_option='USER_ENTERED')
        worksheet.update('B1', right_now_time, value_input_option='USER_ENTERED')
        
        # loops through the Job #, Weight, Direct Hours
        for i,row in enumerate(b[state]):
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
        
        # Removes the in process tag
        worksheet.update('C1', '')
























