# -*- coding: utf-8 -*-
"""
Created on Tue Mar 23 15:33:12 2021

@author: CWilson
"""

import glob
import os
import pandas as pd
import datetime
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
import sys
sys.path.append('c://users//cwilson//documents//python//Weekly Shop Hours Project//')
from TimeClock_Job_Code_Detail import get_most_current_html
from TimeClock_Tools_Employee_Department import download_most_current_employee_department_csv


#%%
def grab_week(state, start, end, create_proof_file, depts_csv):

    #%% Download the state's Job Code Detail File and seperate into direct/indirect
    
    while True:
        try:
            get_most_current_html(state, start, end)
            # Grab all HTML files in downloads
            list_of_files = glob.glob("C://Users//Cwilson//downloads//*.html") # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            state_files = [f for f in list_of_files if state + " Job Code Detail" in f]
            # Get the most recent file for that state
            latest_file = max(state_files, key=os.path.getctime)
            # check to make sure the creation date is today
            file_date = datetime.datetime.fromtimestamp(os.path.getctime(latest_file)).date()
            # if the file date is not today, throw an error
            if file_date != datetime.datetime.today().date():
                2 + "2"
            
            print(latest_file)
            break
        except:
            pass
    

    
    

    
    # converts the html to dataframe
    df = pd.read_html(latest_file)[0] # for some reason it is a list len=1 with the df in it
    # deletes the file from the directory
    os.remove(latest_file)
    # grabs the header row, which has index 0
    header = df.iloc[0]
    # seperates the df from the header
    df = df[1:]
    # makes the header variable the header
    df.columns = header
    # finds the index to sepearte indirect and direct
    df_break = df[df['Job Code'] == "Group STILL ACTIVE Total"].index.tolist()[0]
    
    # Seperate the indirect by the top half of the table. Remove nan rows
    indirect = df[:df_break-2]
    # Seperate direct section as the bottom half. Remove nan rows
    direct = df[df_break+1:-3]
    # reset the index for the direct df
    direct = direct.reset_index(drop=True)
    
    
    
    #%% Categorize the indirect categories
    
    # find the rows in the indirect df that are null based on the 'Number' column
    indirect_na_idx = indirect[indirect['Number'].isnull()].index.tolist()
    # initialize the list for indirect_dfs
    indirect_dfs = []
    for i,idx in enumerate(indirect_na_idx):
        # if it is the first blank row in the df
        if i == 0:
            category_df = indirect[:idx-2]
        # all other rows in the df except the last section
        else:
            category_df = indirect[indirect_na_idx[i-1]:idx-2]
            
        # append to the list
        indirect_dfs.append(category_df)
    # get the last category
    category_df = indirect[indirect_na_idx[-1]:-1]
    indirect_dfs.append(category_df)
    
    
    #%% Categorize the direct categories
    # find the rows in the indirect df that are null based on the 'Number' column
    direct_na_idx = direct[direct['Number'].isnull()].index.tolist()
    
    # initialize the list for indirect_dfs
    direct_dfs = []
    for i,idx in enumerate(direct_na_idx):
        # if it is the first blank row in the df
        if i == 0:
            category_df = direct[:idx-1]
        # all other rows in the df except the last section
        else:
            category_df = direct[direct_na_idx[i-1]+1:idx-1]
        
        # append to the list
        direct_dfs.append(category_df)
    # get the last category
    category_df = direct[direct_na_idx[-1]+1:-1]
    if not category_df.empty:
        direct_dfs.append(category_df)
    
    #%% Converts all the employee numbers and hours to ints and floats respectively

    
    def fix_jobcode_and_to_ints(df_list):
    
        replace_direct_dfs = []
        for df in df_list:
            df1 = df.copy()
            df1['Number'] = df['Number'].astype(int)
            df1['Job Code'] = df.iloc[0,0]
            for col in df.columns[3:]:
                # change the 'Total', 'Regular', 'Overtime 1', 'Overtime 2' columns to floats
                df1[col] = df1[col].astype(float)
            
            replace_direct_dfs.append(df1)
        return replace_direct_dfs
    
    direct_dfs = fix_jobcode_and_to_ints(direct_dfs)
    indirect_dfs = fix_jobcode_and_to_ints(indirect_dfs)    
    
    
    
    
    
    
    #%%  read and process the department
    
    # get file creation timestamp
    employee_info_file_date = os.path.getctime(depts_csv)
    # convert the timestamp to datetime
    employee_info_file_date = datetime.datetime.fromtimestamp(employee_info_file_date)
    # convert datetime to a readable string
    employee_info_file_date = employee_info_file_date.strftime("%m/%d/%Y %I:%H %p")
    print("\n")
    print("Employee Department information from: " + employee_info_file_date)
    print("\n")

    
    # open the employee information csv file
    ei = pd.read_csv(depts_csv)

    # Drop any entries that don't have department filled out
    ei = ei.dropna()
    # unique departments 
    # all_depts = ei['<DEPARTMENT>'].unique()
    # all the maintenance staff
    # maint = ei[ei['<DEPARTMENT>'] == 'MAINTENANCE']
    # need to make a list of what departments are production related
    prod_depts = ['RECEIVING', 'FIT/WELD', 'DRIVER', 'PAINT/LOADING', 'QA/QC',
                  'DETAIL', 'SAW AND DRILL', 'NS-PAINT/LOADING', 'FABRICATION/DRIVER',
                  'NS-FIT/WELD', 'FAB', 'WELD', 'NS-SAW AND DRILL','FITTER',
                  'NS-DETAIL', 'NS-RECEIVING', 'NS-QA/QC', 'NS-CNC OPERATOR',
                  'FIT/WELD-STEEL BAY']
    # Get only the employees categorized by production departments
    prod_employees = ei[ei['<DEPARTMENT>'].isin(prod_depts)]
    # Get just the employee IDs of production employees
    prod_numbers = prod_employees['<NUMBER>'].tolist()
    
    # CSM Shop B Employee Numbers
    shop_b_ids = [2001, 2015, 2029]
    # shop_b_ids = []
    # put other employees that don't get counted in this list
    other_employees_to_delete = []
    employees_to_delete = shop_b_ids + other_employees_to_delete
    
    # Remove each employee from the production IDs
    for employee in employees_to_delete:
        prod_numbers.remove(employee)

    
    #%% Removing and moving things from indirect and direct
    
    
    # job codes that need to be completely removed
    indirect_job_codes_to_delete = ["650-PTO (Shop)", "656-Emergency Paid Leave", 
                            "450-Work for others", "651-Holiday"]
    
    # Job codes which need to be counted as direct
    # Removed "350-Inventory" from the list 3/30/2021
    indirect_job_codes_to_move = ["400-Trucking", "200-Quality Control", 
                                  "250-RECEIVING / STEEL BAY", ]
    
    
    # Jobs that need to be moved to indirect
    direct_job_codes_to_move = ["1626-Federalsburg Plant Improvement",
                                "302002-Trans&Rollover conveyors",
                                "302003-Overton paint shop clean up",
                                "302004-Overton Paint shop cranes"]
    
    # New list of indirect dfs
    indirect_dfs2, direct_dfs2 = [], []
    
    for df in indirect_dfs:
        # don't process if the df is already empty from not having produciton employees
        if df.empty:
            continue
        # grab the job code listed in the df
        job_code = df.iloc[0]["Job Code"]
        # if it is direct, move it then go to next dataframe in list
        if job_code in indirect_job_codes_to_move:
            direct_dfs2.append(df)
            continue
        # append to new list of indirect dfs
        if job_code not in indirect_job_codes_to_delete:
            indirect_dfs2.append(df)

    
    for df in direct_dfs:
        # don't bother with this dataframe if it is empty
        if df.empty:
            continue
        # grab the job code from cell A1
        job_code = df.iloc[0]["Job Code"]
        # move to indirect_dfs2 if it is in the list of jobs to move
        if job_code in direct_job_codes_to_move:
            indirect_dfs2.append(df)
        # if it is not in the jobs to be moved, put it in direct_dfs2
        else:
            direct_dfs2.append(df)
    
    
    shop_b_dfs = []
    
    # remove any non-production employees from indirect hours
    for i, df in enumerate(indirect_dfs2):
        shop_b_only = df[df['Number'].isin(shop_b_ids)]
        if not shop_b_only.empty:
            shop_b_dfs.append(shop_b_only)
        
        prod_only = df[df['Number'].isin(prod_numbers)]        
        indirect_dfs2[i] = prod_only
    
    # remove any non-production employees from direct hours    
    for i, df in enumerate(direct_dfs2):
        shop_b_only = df[df['Number'].isin(shop_b_ids)]
        if not shop_b_only.empty:
            shop_b_dfs.append(shop_b_only)
            
        prod_only = df[df['Number'].isin(prod_numbers)]
        direct_dfs2[i] = prod_only            
    
    
    if state =='TN':
        shop_b_sums = pd.DataFrame(columns=direct_dfs[0].columns)
        for df in shop_b_dfs:
            summed = df.sum()
            shop_b_sums = shop_b_sums.append(summed, ignore_index=True)
        shop_b_sums = shop_b_sums.sum()
        
        print("\n\nSHOP B NUMBERS")
        print('Total:  \t' + str(shop_b_sums['Total']))
        print('Overtime: \t' + str(shop_b_sums['Overtime 1']))
        print('\n')





    #%% Gives a summary of the hours as a digestable csv file
    if create_proof_file:
        # create the dataframe with known columns
        this_df = pd.DataFrame(columns=indirect_dfs2[0].columns)
        # let the first row designate the indirect section
        this_df.loc[0,"Job Code"] = "Indirect"
        # create an empty df that keeps adding up the indirect sums
        sum_of_indirect = pd.DataFrame(dtype=float)
        # append the individual indirect dfs to this_df 
        for df in indirect_dfs2:
            if df.shape[0] == 0:
                continue
            
            this_df = this_df.append(df, ignore_index=True)
            sum_of_df = df.sum()
            sum_of_df['Job Code'] = 'Sum of: ' + df.iloc[0]["Job Code"]
            sum_of_df['Number'] = ''
            sum_of_df['Name'] = ''
            sum_of_indirect = sum_of_indirect.append(sum_of_df, ignore_index=True)
            this_df = this_df.append(sum_of_df, ignore_index=True)
            this_df = this_df.append(pd.Series(dtype=float), ignore_index=True)
        
        # calaulte the sum of each of the hours
        sum_of_indirect = sum_of_indirect.sum()
        # rename the job code 
        sum_of_indirect['Job Code'] = 'Sum of Indirect'
        # append it to the df that becomes a csv
        this_df = this_df.append(sum_of_indirect, ignore_index=True)
        # let the next row be the designator for direct labor
        this_df = this_df.append({"Job Code":"Direct"}, ignore_index=True)
        # create an empty df to add up the direct numbers
        sum_of_direct = pd.DataFrame(dtype=float)
        # append all of the individual direct dfs to this_df
        for df in direct_dfs2:
            if df.shape[0] == 0:
                continue
            
            this_df = this_df.append(df, ignore_index=True) 
            sum_of_df = df.sum()
            sum_of_df['Job Code'] = 'Sum of: ' + df.iloc[0]['Job Code']
            sum_of_df['Number'] = ''
            sum_of_df['Name'] = ''
            sum_of_direct = sum_of_direct.append(sum_of_df, ignore_index=True)
            this_df = this_df.append(sum_of_df, ignore_index=True)
            this_df = this_df.append(pd.Series(dtype=float), ignore_index=True)
        sum_of_direct = sum_of_direct.sum()
        sum_of_direct['Job Code'] = 'Sum of Direct'
        this_df = this_df.append(sum_of_direct, ignore_index=True)
        
        
        # file timestamp
        csv_file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%m-%S")
        # generate the file name for the output
        csv_detail_summary_name = "c://users//cwilson//downloads//" + state + "_" + start.replace("/","-") + "_summary_of_hours " + csv_file_timestamp + ".csv"
        # create a csv file with this_df
        this_df.to_csv(csv_detail_summary_name, index=False)
        print("Created hours summary file: " + csv_detail_summary_name)
    
    #%%
    # empty dfs to grab all the hours in
    indirect_hours = pd.DataFrame(columns=direct.columns[3:], data=[[0,0,0,0]])
    direct_hours = pd.DataFrame(columns=direct.columns[3:], data=[[0,0,0,0]])
    
    # iterate thru indirect dfs and sum
    for df in indirect_dfs2:
        summed = df.sum()
        for col in indirect_hours.columns:
            indirect_hours[col] = indirect_hours[col] + summed[col]
    
    # iterate thru direct dfs and sum
    for df in direct_dfs2:
        summed = df.sum()
        for col in direct_hours.columns:
            direct_hours[col] = direct_hours[col] + summed[col]
    
    
    #%% 
    if state == 'TN':
        sheet = 'CSM QC Form'
    elif state == 'MD':
        sheet = 'FED QC Form'
    else:
        sheet = 'CSF QC Form'
    
    fab_listing = grab_google_sheet(sheet, start, end)

    

    #%%
    final_tons = fab_listing.sum()['Weight'] / 2000
    final_direct_labor = direct_hours['Total']
    final_ttl_labor = direct_hours['Total'] + indirect_hours['Total']
    final_ot_hours = direct_hours['Overtime 1'] + indirect_hours['Overtime 1']
    return [final_tons, final_direct_labor, final_ttl_labor, final_ot_hours]


#%% THIS RUNS THE SCRIPT OVER AND OVER A RANGE OF TIME


def run_range(start_date="03/21/2021", num_weeks=1):
    
    # this gets the most current employee departments csv file name
    while True:
        try:
            download_most_current_employee_department_csv()
            # Grab all HTML files in downloads
            list_of_csvs = glob.glob("C://Users//Cwilson//downloads//*.csv") # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            employe_info_csvs = [f for f in list_of_csvs if "Employee Information" in f]
            # Get the most recent file for that state
            latest_employee_departments = max(employe_info_csvs, key=os.path.getctime)
            # check to make sure the creation date is today
            file_date = datetime.datetime.fromtimestamp(os.path.getctime(latest_employee_departments)).date()
            # if the file date is not today, throw an error
            if file_date != datetime.datetime.today().date():
                2 + "2"            
        
            print('Newest employee department: ', latest_employee_departments)
            break
        except:
            pass
    

    
    
    ''' MAKE SURE THE START DATE IS A SUNDAY '''
    # create a file name with the start date and the number of weeks ran
    file_output_name = "Weekly report starting "+ start_date.replace("/", "-") + " num weeks-" + str(num_weeks)
    file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%m-%S")
    # make the file go to the downloads folder
    file_name = "c:\\users\\cwilson\\downloads\\" + file_output_name + " " + file_timestamp + ".csv"    
    # convert the start date to a datetime
    dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
    # I don't care about the time, just the date and this is easier output for the files
    user_start = datetime.date(dt.year, dt.month, dt.day)
    # initialize the dataframe
    state_df = pd.DataFrame(columns=['Date', 'State', 'Tons', 'Direct Labor', 'TTL Labor', 'OT Hrs'])
    # create a counter to iterate thru the df
    idx = 0
    # loop thru the 3 shops
    for state in ['DE', 'MD', 'TN']:
        # have to reassign the start before the state changes
        start = user_start
        # loop through as many weeks as were defined
        for week in range(0,num_weeks):
            # end is 6 days after start so its sunday to saturday
            end = start + datetime.timedelta(days=6)
            # convert startdate to a string
            start_str = start.strftime("%m/%d/%Y")
            # convert end date to a string
            end_str = end.strftime("%m/%d/%Y")
            # run the bigass function to get hours
            hours = grab_week(state, start_str, end_str, create_proof_file=True, depts_csv=latest_employee_departments)
            # assign the time to be the start date
            state_df.loc[idx, 'Date'] = start
            # state
            state_df.loc[idx, 'State'] = state
            # tons is the first value in the list
            state_df.loc[idx, 'Tons'] = hours[0]
            # For the following, the [0] is because they are passed as pd.series and I am lazy
            state_df.loc[idx, 'Direct Labor'] = hours[1][0]
            state_df.loc[idx, 'TTL Labor'] = hours[2][0]
            state_df.loc[idx, 'OT Hrs'] = hours[3][0]
            
            print(state, start_str, end_str)
            # print(state_df)
            # increase the index row by one
            idx += 1
            # adjust the start date for when grabbing multiple weeks
            start = start + datetime.timedelta(days=7)
        
    # print a copy of the df 
    print(state_df)
    print('\n', file_name)
    # write the df to a csv file
    state_df.to_csv(file_name)
    
    # delete the employee departments csv once all done
    os.remove(latest_employee_departments)


run_range(start_date = "05/02/2021", num_weeks = 1)
print("V:\PRODUCTION REPORT FROM US\Productivity Reports\Prod Reports 2021")


