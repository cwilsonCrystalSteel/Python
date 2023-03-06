# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 14:02:51 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import datetime
import timeit
from Combine_HTML_and_SHEET import return_combined_hours_weights

def gather_states_data(state, start_date, end_date):

    # wrapped the whole thing in a try-except in an attempt to limit failed attempts
    # Also logs the errors 
    try:
        start_time = timeit.default_timer()
        
        print('Started ' + state)

        # get todays date in string %m/%d/%y format to run timeclock on 
        # today = datetime.date.today().strftime("%m/%d/%Y")
        # Grabs the latest data from the Daily Fab Listing
        # Only takes the last 100 entries and then takes only todays work

        # Returns the direct labor df, date of update, time of update
        combined_data = return_combined_hours_weights(state, start_date, end_date)
        # The first item in the list is todays direct hours
        df = combined_data[0]
        # I want to sort the data by putting the jobs with the most weight in column 1
        df = df.transpose()
        # after transposing, sort by the weight and then by the # of hours worked
        df = df.sort_values(by=['Weight','Hours'], ascending=False)
        # transpose it back to normal shape so that it can be grabbed as lists like normally
        df = df.transpose()
        
        # First row of data - job #'s
        row1 = df.columns.tolist()
        # Second row of data - weights on that job
        row2 = df.loc['Weight'].tolist()
        # Third row of data - hours on that job
        row3 = df.loc['Hours'].tolist()
        # fourth row of data - eva model hours
        row4 = df.loc['EVA Hours'].tolist()
        
        print('Finished: ' + state)
        elapsed = timeit.default_timer() - start_time
        print(elapsed)
        
        # returns the 4 rows that are used to update the behind the scenes files
        return [row1, row2, row3, row4]
    
    except Exception as e:
        # error log directory
        error_log = "C:\\Users\\cwilson\\Documents\\Python\\Publish to Live\\Error_logs\\"
        # gets current date to timestamp the file
        error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        # create the entire file name
        file_name = error_log + "gather_data_error-" + error_date + ".txt"
        # open the file
        file = open(file_name, 'w')
        # write the state name to the file
        file.write(state + '\n')
        # write the error to the file
        file.write(str(e))
        # close the file
        file.close()
        



























