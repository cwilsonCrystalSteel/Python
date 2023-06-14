# -*- coding: utf-8 -*-
"""
Created on Mon Mar 15 15:20:52 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import time
import gspread
import pandas as pd
import datetime
from production_dashboards_google_credentials import init_google_sheet

json_file = 'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-e051ae12d1ef.json'

history_sheets = ['CSM History', 'CSF History', 'FED History']

headers = ['Date', 'Tons', 'Earned', 'Actual', 'Efficiency']

try:
    
    for sheet_name in history_sheets:
        
        # Google sheets ID Number
        live_output_google_sheets_key = "1RZKV2-jt5YOFJNKM8EJMnmAmgRM1LnA9R2-Yws2XQEs"
        # Grab the spreadsheet
        sh = init_google_sheet(live_output_google_sheets_key, json_file)
        # Grab the individual sheet
        worksheet = sh.worksheet(sheet_name)
        
        # Get the current datetime as a string for file naming
        date_now = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")
        # create the file name with global destination
        file_name = "C:\\users\\cwilson\\documents\\python\\Production_Dashboard_temp_files\\Archived History\\" 
        file_name += "Weekly " + sheet_name + " " + date_now + ".csv"
        # convert the current history sheet to a df
        df = pd.DataFrame(worksheet.get_all_records())
        # create the csv file from the df
        df.to_csv(file_name)
        
        worksheet.clear()
        
        # put the column names back into the sheet
        for i,val in enumerate(headers):
            time.sleep(1.1)
            cell = chr(i+97) + "1"
            worksheet.update(cell, val, value_input_option='USER_ENTERED')
            
except Exception as e:
    # error log directory
    error_log = "C:\\Users\\cwilson\\Documents\\Python\\Production_Dashboard_temp_files\\Errors\\"
    # gets current date to timestamp the file
    error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # create the entire file name
    file_name = error_log + "archive_error-" + error_date + ".txt"
    # open the file
    file = open(file_name, 'w')
    # write the state name to the file
    file.write(sheet_name + '\n')
    # write the error to the file
    file.write(str(e))
    # close the file
    file.close()