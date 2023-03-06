# -*- coding: utf-8 -*-
"""
Created on Fri Apr  9 12:11:12 2021

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import pandas as pd
from bs4 import BeautifulSoup
import datetime
import glob
import os
sys.path.append('c://users//cwilson//documents//python//Attendance Project//')
sys.path.append('c://users//cwilson//documents//python//')

from TimeClock_Payroll_Weekly_Punch import download_most_current_weekly_punch_html
from TimeClock_Tools_Employee_Location import download_most_current_employee_location_csv
from attendance_google_sheets_credentials_startup import init_google_sheet


#%% These two loops get all of the data needed for the script to run

# get the current timestamp
now = datetime.datetime.now()
# get the date as a string for TimeClock stuff
todays_datestring = now.strftime("%m/%d/%Y")

# Use this while-try loop to download the html time-in data and retrieve the file from the downloads folder
while True:
    try:
        download_most_current_weekly_punch_html(todays_datestring)
        # Grab all HTML files in downloads
        list_of_htmls = glob.glob("C://Users//Cwilson//downloads//*.html") # * means all if need specific format then *.csv
        # Create a list with only the states we want to look at
        weekly_punch_files = [f for f in list_of_htmls if "Weekly Punch" in f]
        # Get the most recent file for that state
        latest_html = max(weekly_punch_files, key=os.path.getctime)
        print(latest_html)
        break
    except:
        pass


# use this while-try loop to download the csv employee location data and retrieve from downloads folder
while True:
    try:
        download_most_current_employee_location_csv()
        # Grab all HTML files in downloads
        list_of_csvs = glob.glob("C://Users//Cwilson//downloads//*.csv") # * means all if need specific format then *.csv
        # Create a list with only the states we want to look at
        employe_info_csvs = [f for f in list_of_csvs if "Employee Information" in f]
        # Get the most recent file for that state
        latest_csv = max(employe_info_csvs, key=os.path.getctime)
        print(latest_csv)
        break
    except:
        pass



#%% I Dont know what I did but this gets all the rows of the html table

data = []
with open(latest_html, 'r') as f:
    
    soup = BeautifulSoup(f, 'html.parser')
    # print(soup.prettify())
    table = soup.find('body')
    rows = table.find_all('tr')

    
    for row in rows:
        
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        cols = [ele for ele in cols if ele]
        # if len(cols) > 2:
        data.append(cols)

# Deletes the HTML file from the downlaods folder
os.remove(latest_html)

#%%
# only keeps lists that are non-empty
new_data = []
for l in data:
    if l:
        new_data.append(l)

#%%
# If the first element in the list represents an integer, it is added to a new list
# basically, if the first element is their employee number then keep it
newer_data = []
for l in new_data:
    if len(l) > 1:
        try:
            int(l[0])
            newer_data.append(l)
        except:
            pass

#%%
# grabbing the employee number, name, time in 
# Not all of the lists have the same length
# some have the Clock date/time & actual date/time
# others have 'missed' in the [2] spot

missed_list = []
in_list = []
actual_time_in_list = []
present_list = []

for l in newer_data:
    # grabs the employee number as the first value in the list
    employee_number = int(l[0])
    # gets the employees name as the second value in the list
    employee_name = l[1]
    # checks what the third value in the list is 
    two_spot = l[2]
    # if the report shows the third item as 'Missed', it moves that employees information to the missed_list
    if two_spot == 'Missed':
        missed_list.append([employee_number, employee_name, two_spot])
        continue
    
    # checks if the fifth item is 'IN' or if it is a date
    four_spot = l[4]
    # if the fifth item is 'IN', the employees time-in will be the fourth item 
    if four_spot == 'IN':
        # l[3] is the time in if it makes it to here
        in_list.append([employee_number, employee_name, l[3]])
        present_list.append([employee_number, employee_name, l[3]])
        
        continue
    elif 'IN' not in l:
        
        # empty list to get the indexes if the value can be construed as a dat
        full_date_idxs = []
        for i, val in enumerate(l):
            try:
                # checks if it is in the form %m/%d/%Y
                datetime.datetime.strptime(val, "%m/%d/%Y")
                # appends the index to the list
                full_date_idxs.append(i)
            except:
                pass
        
        # if the list has multiple dates in it (should have either 1 or 2 dates)
        # it will grab the time as the value after the second date
        if len(full_date_idxs) > 1:
            time_in = l[full_date_idxs[1]+1]
        # if the list is only 1 long, it will take the time from the one after the date
        elif len(full_date_idxs) == 1:
            time_in = l[full_date_idxs[0]+1]
            

        actual_time_in_list.append([employee_number, employee_name, time_in])
        present_list.append([employee_number, employee_name, time_in])
        continue
    
    else:
        # l[5] is the time-in if it makes it this far
        actual_time_in_list.append([employee_number, employee_name, l[5]])
        present_list.append([employee_number, employee_name, l[5]])

        
present_df = pd.DataFrame(columns=['Number', 'Name', 'Time in'], data=present_list)
# present_df['Number'] = present_df['Number'].astype(int)
present_df['Time in'] = pd.to_datetime(present_df['Time in'])
present_df['Time in'] = [time.time() for time in present_df['Time in']]

missed_df = pd.DataFrame(columns=['Number', 'Name', 'Time in'], data=missed_list)

#%%
# Need to get the most current active employee list
# compare the active employee list to those clocked in
# get rid of all employees that clocked in on time
# move the employees that are not clocked in to a list
# move the late employees to another list

employee_locations = pd.read_csv(latest_csv)
# Delete the Employee Information CSV file from the downloads folder
os.remove(latest_csv)
# rename the columns b/c it makes it easier to work with
employee_locations = employee_locations.rename(columns={'<NUMBER>':'ID',
                                                 '<FIRSTNAME>':'First',
                                                 '<LASTNAME>':'Last',
                                                 '<LOCATION>':'Location',
                                                 '<CLASS>':'Shift',
                                                 '<SCHEDULEGROUP>':'Productive'})
# create a 'Time in' column with a palceholder time of 23:59
employee_locations['Time in'] = datetime.time(23,59)

# iterate through the list of those that are clocked in 
for person in present_list:
    # the employee number is the first item in the list
    num = person[0]
    # the time-in is the 3rd item in the list, converst it from string to datetime.time
    timein = datetime.datetime.strptime(person[2], "%H:%M %p").time()
    # Find the index in the employee_locations dataframe based on employee ID number
    idx = employee_locations.index[employee_locations['ID'] == num].tolist()[0]
    # sets the 'Time in' column from 23:59 to the actual time they clocked in
    employee_locations.loc[idx,'Time in'] = timein

# sort the data frame by Location, then by the employee ID number
# this makes it easier to post to the google sheet
employee_locations = employee_locations.sort_values(by=['Location','ID'])
# drop employees from the list if the Location column is na
employee_locations = employee_locations[~employee_locations['Location'].isna()]
# dropping KRS employees from the dataframe for simplicity sake
employee_locations = employee_locations[~employee_locations['Location'].str.contains('KRS')]




''' THIS IS A FILLER UNTIL THE SHIFTS ARE UPDATED FOR ALL EMPLOYEES '''
employee_locations['Shift'] = employee_locations['Shift'].fillna(999)
''' THIS IS A FILLER UNTIL THE SHIFTS ARE UPDATED FOR ALL EMPLOYEES '''


''' I think this is what I will need to do once the shifts classification is filled out '''
'''
# get the time to determine if running 1st or 2nd shift attendance
time_now = now.time().replace(second=0,  microsecond=0)

shift_to_run = 1

if time_now.hour <= 12:
    shift_to_run = 1
else:
    shift_to_run = 2

employee_locations = employee_locations[employee_locations['Shift'] == shift_to_run]
'''
''' I think this is what I will need to do once the shifts classification is filled out '''



#%%

# get the start of the shift time from the 'Start of Shift Times' sheet
sh = init_google_sheet("1zyKqNfHnjhrzW_rgqXG4Dk9CV53kfg9TaJj26yP1Ytg")
worksheet = sh.worksheet('Start of Shift Times')
worksheeet_list = worksheet.get_all_values()





# absent if you clock in after 9:30
absent_df = employee_locations[employee_locations['Time in'] > datetime.time(9,30)]
# convert to a list of lists - makes it easier to post to google sheet
absent_list = absent_df.values.tolist()





# late if you clock in after 6 am
late_df = employee_locations[employee_locations['Time in'] > datetime.time(6,0)]
# remove anybody from the late_df if their time is 23:59
late_df = late_df[late_df['Time in'] != datetime.time(23,59)]

# convert to a list
late_list = late_df.values.tolist()




# locations = pd.unique(employee_locations['Location'])




#%% Perform operations on the LATE sheet 





def update_sheet_return_old_df(the_list, headers, sheet_name):
    import time
    
    try:
        # go to the google sheeet
        sh = init_google_sheet("1zyKqNfHnjhrzW_rgqXG4Dk9CV53kfg9TaJj26yP1Ytg")
        # go to the late tab
        worksheet = sh.worksheet(sheet_name)
        # grab all the values of yesterdays worksheet
        worksheet_list = worksheet.get_all_values()
        # create a dataframe from the sheet
        old_df = pd.DataFrame(worksheet_list[1:], columns=worksheet_list[0])
        # clears the worksheet for today's data
        worksheet.clear()
        
        # create the headers in the worksheet
        for i,header in enumerate(headers):
            # get the cell letter with chr(i+97)
            cell = chr(i+97) + '1'
            # update the sheet with the header value
            worksheet.update(cell, header)
            time.sleep(1.1)
        
        # create a comments column for managers 
        cell = chr(i + 1 + 97) + '1'
        worksheet.update(cell, 'Comments for: ' + datetime.date.today().strftime("%m/%d/%Y"))
        time.sleep(1.1)
        
      
        # iterate through the lists of lists
        for i,l in  enumerate(the_list):
            # iterate through the individual employee's list
            for j,val in enumerate(l):
                # get the column letter
                col_letter = chr(j+97)
                # get the row number and convert to string
                # The +2 goes to a regular index and then skps the first row 
                row_number = str(i+2)
                # create the cell number
                cell = col_letter + row_number
                # if it is the last value in the employee list, it is a time
                if j+1 == len(l):
                    # convert that time to a string to easily put in google sheet
                    val = val.strftime("%H:%M %p")
                
                # update the cell 
                worksheet.update(cell, val, value_input_option='USER_ENTERED')
                time.sleep(1.3)

    
    except Exception as e:
        # error log directory
        error_log = "C:\\Users\\cwilson\\Documents\\Python\\Attendance Project\\Error_logs\\"
        # gets current date to timestamp the file
        error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        # create the entire file name
        file_name = error_log + "Attendance Error-" + error_date + ".txt"
        # open the file
        file = open(file_name, 'w')
        # write the error to the file
        file.write(str(e))
        # close the file
        file.close()
        
        
    return old_df






def create_file_name(string, df):
    get_old_date = df.columns.tolist()[-1][-10:]
    get_old_date = get_old_date.replace('/','-')
    
    base_history = "c://users//cwilson//documents//python//Attendance Project//History//"
    file_name = base_history + string + " " + get_old_date + ".csv"
    return file_name


old_late = update_sheet_return_old_df(late_list, late_df.columns.tolist(), 'Late')
old_late.to_csv(create_file_name('Late', old_late))

old_absent = update_sheet_return_old_df(absent_list, absent_df.columns.tolist(), 'Absent')
old_absent.to_csv(create_file_name('Absent', old_absent))








