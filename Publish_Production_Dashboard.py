# -*- coding: utf-8 -*-
"""
Created on Fri Mar 19 09:50:48 2021

@author: CWilson
"""


import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import gspread
from production_dashboards_google_credentials import init_google_sheet
import time
import datetime
from CHANGE_THIS_NAME import download_data, get_production_dashboard_data
import pandas as pd


def publish_dashboard(dashboard_name):
    download_base = "C:\\users\\cwilson\\documents\\python\\Production_Dashboard_temp_files\\"
    
    today = datetime.date.today()
    
    dashboard_options = {}
    
    dashboard_options['Yearly'] = {'sheet_key':'1bEzy9UZdUrvdesdxMBum2NIHSm0ZooBYr3_5jLsjeOU',
                                   'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-890ed2bf828b.json',
                                   'download_folder': download_base + "Yearly\\",
                                   'start_date':datetime.date(today.year, 1, 1).strftime('%m/%d/%Y'),
                                   'end_date':datetime.date(today.year, 12, 31).strftime('%m/%d/%Y')}
    
    dashboard_options['Monthly'] = {'sheet_key':'1kwuWsOEEPcJWfl2EBkhOaSxAGl1vtPNtqoqIvsWjxmc',
                                    'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-890ed2bf828b.json',
                                    'download_folder': download_base + "Monthly\\",
                                    'start_date':datetime.date(today.year, today.month, 1).strftime('%m/%d/%Y'),
                                    'end_date':(datetime.date(today.year, today.month+1, 1) + datetime.timedelta(days=-1)).strftime('%m/%d/%Y')}
    
    dashboard_options['Weekly'] = {'sheet_key':'1RZKV2-jt5YOFJNKM8EJMnmAmgRM1LnA9R2-Yws2XQEs',
                                   'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-890ed2bf828b.json',
                                   'download_folder': download_base + "Weekly\\",
                                   'start_date':(today - datetime.timedelta(days=today.weekday() + 1)).strftime('%m/%d/%Y'),
                                   'end_date':(today - datetime.timedelta(days=today.weekday() + 1 - 6)).strftime('%m/%d/%Y')}
    
    dashboard_options['Daily'] = {'sheet_key':'1bpb75pCrsRh7t4FZr1bMILCRzpVGyiOCws4oCh2nC5c',
                                  'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-daily-1568a4c99f1f.json',
                                  'download_folder': download_base + "Daily\\",
                                  'start_date':today.strftime('%m/%d/%Y'),
                                  'end_date':today.strftime('%m/%d/%Y')}

    
    key = dashboard_options[dashboard_name]['sheet_key']
    json_file = dashboard_options[dashboard_name]['json_file']
    download_folder = dashboard_options[dashboard_name]['download_folder']
    start_date = dashboard_options[dashboard_name]['start_date']
    end_date = dashboard_options[dashboard_name]['end_date']
    
    
    states = ['TN', 'MD', 'DE']
    
    
    try:
        base_data = download_data(start_date, end_date, download_folder)
        state_data = get_production_dashboard_data(start_date, end_date, base_data)
    except Exception as e:        
        # create a copy of the current variables in memory
        current_variables = dir().copy()
        # error log directory
        error_log = "C:\\Users\\cwilson\\Documents\\Python\\Production_Dashboard_temp_files\\Errors\\"
        # gets current date to timestamp the file
        error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        # create the entire file name
        file_name = error_log + 'Gathering Data Error (' + dashboard_name + ') ' + error_date + ".txt"
        # open the file
        file = open(file_name, 'w')
        # write the dahsboard name to the file
        file.write(dashboard_name + '\n')
        # write the error to the file
        file.write(str(e))
        file.write('\n--- Start of the Variables ---\n')
        # iterate thru each variable 
        for var_name in current_variables:
            file.write('\n' + var_name + '\n')    

    
    
    
    ''' This snippet makes it so either all states update or none update '''
    # defaults to be allowed to run
    good_to_run = True
    # checks if the value for each state is a list or not
    # if there was an error during the 'gather_states_data()', the value becomes a NoneType
    for val in state_data:
        if isinstance(state_data[val], pd.DataFrame):
            if state_data[val].shape[0] == 0:
                print(val, " is bad")
                good_to_run = False
                break
        else:
            print(val, " is bad")
            # toggles the ability to run the program if there is a NoneType present
            good_to_run = False
            break
    
    
    print(dashboard_name)
    try:
        
        if good_to_run:
            # counts the number of elements that will be updated by looping thru all the lists
            count = 0
            for state in states:
                count += state_data[state].count().sum()
            print('# of cells to be updated: ' + str(count))
            # use count to determine the sleep timer for API limits
            # setting sleep timer to 1.5 second for now b/c I don't want to figure out how it needs to be changed
            x = 1.1
            
            # the ID # for the Live sheet
            # live_output_google_sheets_key = "12yFpSXyblbhueEM6e5vz_WaEXJKTsozHvkfS9gjux3w"
            # Grab the spreadsheet
            sh = init_google_sheet(key, json_file)
            # the individual sheet is grabbed during the loop below
            
            # have the right now time calculation outside the loop so all shops 
            # get updated with the same timestamp
            right_now_dt = datetime.datetime.now()
            # break it into date string
            right_now_date = right_now_dt.date().strftime("%m/%d/%Y")
            # break it into time string
            right_now_time = right_now_dt.time().strftime("%I:%M %p")
            # have right now be a string
            right_now = right_now_dt.strftime('%m/%d/%Y %I:%M %p')
            
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
                # totals_lists = totals(base_sheet_name + ' QC Form')       
                
                # open the specific shop's behind-the-scenes sheet
                worksheet = sh.worksheet(base_sheet_name)
                print("Editting " + base_sheet_name + " behind the scenes")
                
                # Grab the date from A1
                last_date = worksheet.acell('A1').value
                # # Grab the Time from B1
                last_time = worksheet.acell('B1').value
                
                
                # convert the last timestamp to a datetime
                last_dt = datetime.datetime.strptime(last_date + ' ' + last_time,
                                                     "%m/%d/%Y %H:%M %p")
                # add 10 minutes to the last datetime
                last_dt_plus_time = last_dt + datetime.timedelta(minutes=10)
                # if it is currently within 10 minutes of the last datetime, break 
                if last_dt_plus_time > right_now_dt:
                    print('Stopping the publishing b/c it has been less than 10 minutes')
                    break
                


                
                
                
                # clears the behind-the-scenes worksheet
                worksheet.clear()
                # Adds the IN PROCESS tag with a timestamp
                worksheet.update('C1', "IN PROCESS as of: " + right_now_time)       
                # pause
                time.sleep(x)
                # Also need to update cells A1 for the Date
                worksheet.update('A1', right_now_date, value_input_option='USER_ENTERED')
                # pause
                time.sleep(x)
                # Update cell B1 with the time
                worksheet.update('B1', right_now_time, value_input_option='USER_ENTERED')
                time.sleep(x)
                worksheet.update('A3', 'Job')
                time.sleep(x)
                worksheet.update('B3', 'Weight')
                time.sleep(x)
                worksheet.update('C3', 'Actual Hours') # the direct hours worked
                time.sleep(x)
                worksheet.update('D3', 'Sold Hours') # the model / EVA hours
                
                
                states_df = state_data[state].reset_index(drop=False).astype(float)
                
                # loops thru each row of the dataframe
                for i in range(0, states_df.shape[0]):
                    # loops thru each column of the dataframe
                    for j in range(0, states_df.shape[1]):
                        # pause for api limits
                        time.sleep(x)
                        # covert column number to letter
                        col = chr(j + 97)
                        # get the row number: +4 so it goes on the 3rd row
                        row = i + 4
                        # create the cell name
                        cell = col + str(row)
                        # get the value, which is just row,col of the df
                        value = states_df.iloc[i,j]
                        # update the worksheet
                        worksheet.update(cell, value)
                        

                # Removes the in process tag
                worksheet.update('C1', '')
                
                # get the summary values for the HISTORY tabs
                states_df_summed = states_df.sum()
                tonnage = states_df_summed['Weight'] / 2000
                earned_hours = states_df_summed['Earned Hours']
                worked_hours = states_df_summed['Worked Hours']
                

                print('now editing the history sheet for ' + state)
                # change to the history sheet
                hist_ws = sh.worksheet(base_sheet_name + ' History')
                # get all of the values as a list of lists
                history = hist_ws.get_all_values()
                #get the number of rows in the history sheet as a str
                row = str(len(history) + 1)
                # if the sheet was cleared, set the starting row as 2
                if row == '1':
                    row == '2'
                
                # combine all the values that are going to be put into the newest row of the sheet
                history_sheet_new_row = [right_now_date + ' ' + right_now_time,
                                         tonnage,
                                         earned_hours, 
                                         worked_hours,
                                         earned_hours / worked_hours]
                
                # loop thru each of the values that are getting pasted into the history
                for i, val in enumerate(history_sheet_new_row):
                    # pause
                    time.sleep(1)
                    # get the column letter
                    col = chr(i+97)
                    # combine col and row to create a cell
                    cell = col + row
                    # update the cell
                    hist_ws.update(cell, val, value_input_option='USER_ENTERED')                

                
                
                
                
                
                
                
                
                
                
    except Exception as e:
        # create a copy of the current variables in memory
        current_variables = dir().copy()
        # error log directory
        error_log = "C:\\Users\\cwilson\\Documents\\Python\\Production_Dashboard_temp_files\\Errors\\"
        # gets current date to timestamp the file
        error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        # create the entire file name
        file_name = error_log + 'Publishing Error (' + dashboard_name + ') ' + error_date + ".txt"
        # open the file
        file = open(file_name, 'w')
        # write the state name to the file
        file.write(state + '\n')
        # write the dahsboard name to the file
        file.write(dashboard_name + '\n')
        # write the error to the file
        file.write(str(e))
        file.write('\n--- Start of the Variables ---\n')
        # iterate thru each variable 
        for var_name in current_variables:
            file.write('\n' + var_name + '\n')
 
    
        # close the file
        file.close()
