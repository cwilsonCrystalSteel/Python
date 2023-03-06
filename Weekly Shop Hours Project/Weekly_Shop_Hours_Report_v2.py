# -*- coding: utf-8 -*-
"""
Created on Wed May  5 10:48:00 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('c://users//cwilson//documents//python//')
import subprocess
import os
import pandas as pd
import datetime
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports






start_date = input("Enter Sunday's date (mm/dd/yyyy): ")

try:
    start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
except:
    print('Invalid Date Format - Please close & try again')
    quit()


# # manually enter when the starting sunday is
# start_date = "08/15/2021"
# # convert the start date string to a datetime
# start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")




''' This json file has the changes needed to be made regarding job codes '''
import json
code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))


# setup the initial directory to put the proof files into
base_directory = "c://users//cwilson//documents//Productive_Hours_Weekly_Report//weekly_shop_hours_dump//"
# get todays date as a string
today_str = datetime.datetime.today().strftime('%m-%d-%Y')
# create new directory name based on todays date
directory_for_dumping = base_directory + today_str + '//'
# check if the directory already exists
if not os.path.exists(directory_for_dumping):
    # if it does not exist, make the directory
    os.makedirs(directory_for_dumping)






# initialize the summary dataframes
summary_de = pd.DataFrame(columns=['Date','Tons','Direct','Total','OT'])
summary_md = pd.DataFrame(columns=['Date','Tons','Direct','Total','OT'])
summary_tn = pd.DataFrame(columns=['Date','Tons','Direct','Total','OT'])


for week in range(0,1):
    print('')
    
    end_dt = start_dt + datetime.timedelta(days=6)
    end_date = end_dt.strftime("%m/%d/%Y")
    
    
    
    
    # download shit from TimeClock & do basic preprocessing
    basis = get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=False)
    

    
    
    
    # get the clock details dataframe
    hours = basis['Clocks Dataframe']
    # get the employee information dataframe
    ei = basis['Employee Information']
    
    # jeff davis
    # ei.loc[152,'Productive'] = 'TN PRODUCTIVE'
    # cornelius rivers
    # ei.loc[160,'Productive'] = 'TN PRODUCTIVE'
    # mike lewis
    # ei.loc[204,'Productive'] = 'TN NON PRODUCTIVE'
    
    ''' Drop the SHop B guys '''
    ei = ei[~ei['Name'].isin(['Robert Burrell', 'Rodger Grant', 'Robert Richardson'])]
    # Only keep the productive guys
    ei = ei[~ei['Productive'].str.contains('NON')]
    # set the index to be the names
    ei = ei.set_index(ei['Name'])
    # get rid of the default location column
    ei = ei.drop(columns=['Location'])
    # set the location column in ei as the start of the Productive tag
    ei['Location'] = ei['Productive'].str[:2]
    
    # set the index to be the names
    hours = hours.set_index(hours['Name'])
    # join the location column to the hours df
    hours = hours.join(ei['Location'])
    # get rid of non productive employees which now have nan for productive
    hours = hours[~hours['Location'].isna()]
    # for some reason after this join there are some people that get double hours
    hours = hours.drop_duplicates()
    # reset the hours index
    hours = hours.reset_index(drop=True)
    # replace the 'no cost code' string with blanks
    hours['Cost Code'] = hours['Cost Code'].replace('no cost code','')
    
    #%% This is what gets the direct and indirect horus into 2 dataframes
    
    # create the combination column with a '<>' to seperate
    # hours['JCC'] = hours['Job #'].astype(str) + '<>' + hours['Cost Code']
    
    
    # delete the stuff like pto
    hours = hours[~hours['Job Code'].isin(code_changes['Delete Job Codes'])]
    # create the indirect df starting with indirect cost codes
    indirect = hours[hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # remove those indirect cost code items from the hours df now
    hours = hours[~hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # append the indirect job codes
    indirect = indirect.append(hours[hours['Job Code'].isin(code_changes['Indirect Job Codes'])])
    # remove those from the hours df now
    hours = hours[~hours['Job Code'].isin(code_changes['Indirect Job Codes'])]
    # start with the things that appear to be indirect but are actually direct
    direct = hours[hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # remove those from the hours df now
    hours = hours[~hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # get the 3 digit items from hours df and move to indirect
    three_digit_indirects = hours[hours['Job #'] < 1000]
    # append the 3digit jobs that are indirect into indirect df
    indirect = indirect.append(three_digit_indirects)
    # delete the 3digit jobs from hours
    hours = hours[~hours.index.isin(three_digit_indirects.index)]
    # append what is left of hours df to direct
    direct = direct.append(hours)
    # sort the direct dataframe
    direct = direct.sort_values(['Name','Job #','Cost Code'])
    # sort the indirect dataframe
    indirect = indirect.sort_values(['Name','Job #','Cost Code'])    
    
    
    # %% group the direct & indirect to remove duplicates of job & name combinations
    # group the direct hours by name and job code
    direct2 = direct.groupby(['Name','Job #','Job Code']).sum()
    # add a is direct = true column
    direct2['Is Direct'] = True
    # create an empty cost code b/c i dont care about cost code for direct work
    direct2['Cost Code'] = ''
    # group the indirect hours by name, job code and cost code b/c cost code is important for 
    # ticket work, revisions, trucking, etc
    indirect2 = indirect.groupby(['Name','Job #','Job Code','Cost Code']).sum()
    # add a is direct = false column
    indirect2['Is Direct'] = False
    # append the dataframes back together 
    all_df = direct2.reset_index(drop=False).append(indirect2.reset_index(drop=False))
    # get the locations of the employees from the location column of ei
    all_df = all_df.join(ei['Location'], on=['Name'])

    # sort by Direct/Indirect, then job , then location, then by name
    all_df = all_df.sort_values(by=['Is Direct','Job #','Location','Name'])    
    
    
    
    

    
    #%% This calculates the overtime hours for each shop
    
    # Sum up the hours worked by each employee
    hour_count = pd.DataFrame(all_df.groupby('Name')['Hours'].sum())
    # get the employees with over 40 hours
    ot_employees = hour_count[hour_count['Hours'] > 40]
    # create a column for the regular hours
    hour_count['Regular'] = hour_count['Hours']
    # overtime starts with regular minus 40 hours
    hour_count['OT'] = hour_count['Hours'] - 40
    # if the OT is negative, get a boolean series
    neg_ot = hour_count['OT'] <= 0
    # multiply the OT column by the inverse of the negative count
    # if it is a negative it gets multiplied by zero, becoming zero
    # the .abs() is b/c some values were negative zero & i dont care for that
    hour_count['OT'] = (hour_count['OT'] * ~neg_ot).abs()
    # Subtract the OT from the Regular (Which is still just the total hours)
    hour_count['Regular'] = hour_count['Regular'] - hour_count['OT']
    # join the employee locations to the dataframe
    hour_count = hour_count.join(ei['Productive'])
    # create a series with the OT hours for each shop
    ot = hour_count.groupby('Productive')['OT'].sum()
    # redo the index names so they are just the state initials
    ot = ot.rename(index = {x:x[:2] for x in ot.index})
    

    
    #%% divides the all_df into portions by location
    
    # create empty dict for each state to have
    states_dict = {}
    # get the states available in the all-df
    states = pd.unique(all_df['Location'])
    
    for state in states:
        # create empty dict for that state
        states_dict[state] = {}
        # get only that states data from the big dataframe
        states_all = all_df[all_df['Location'] == state]
        # get the direct dataframe
        states_direct = states_all[states_all['Is Direct'] == True]
        # get the indirect dataframe
        states_indirect = states_all[states_all['Is Direct'] == False]
        # apply the direct dataframe to the dict
        states_dict[state]['Direct'] = states_direct
        # apply the indirect dataframe to the dict
        states_dict[state]['Indirect'] = states_indirect


    #%% This is to get the direct hours grouped by job for the productin worksheet
    # only get direct from all_df -> because it already has the location present
    # group by the job # and by location
   
    # reset the index 
    direct3 = all_df[all_df['Is Direct']].groupby(['Job #', 'Location']).sum()
    # drop the is direct column
    direct3 = direct3.drop(columns=['Is Direct'])
    # sort by location then by job #
    direct3 = direct3.sort_values(by=['Location','Job #'])
    # reset the index bc i need those indexes as columns
    direct3 = direct3.reset_index(drop=False)
    # create a new df that will be the csv output
    production_worksheet = pd.DataFrame(columns=['Job #', 'Hours'])
    # loop thru each state
    for state in states:
        # get that states data
        chunk = direct3[direct3['Location'] == state]
        # append a row as the state name
        production_worksheet = production_worksheet.append({'Job #':state}, ignore_index=True)
        # append the chunk df
        production_worksheet = production_worksheet.append(chunk[['Job #','Hours']])
        # append the total for that shop so I dont have to click around to other places
        production_worksheet = production_worksheet.append({'Job #':'Total', 'Hours':chunk.sum()['Hours']}, ignore_index=True)
    
    # write to a csv file in the same directory as the other output files go to 
    production_worksheet.to_csv(directory_for_dumping+'Production Worksheet Output.csv', index=False)
    
            
        
    
    #%% this is where I combine the indirect & direct into a CSV readable format with sums & shit
    
    # initialize the summary dataframe which gets printed for easy access
    summary = pd.DataFrame(columns=['Tons','Direct', 'Indirect','Total','OT'])
    
    for state in states:
        # sheet names for grabbing data from fablisting
        if state == 'TN':
            sheet = 'CSM QC Form'
        elif state == 'MD':
            sheet = 'FED QC Form'
        elif state == 'DE':
            sheet = 'CSF QC Form'
        
        # grab that states indirect dataframe
        df = states_dict[state]['Indirect'].copy()
        # only keep certain columns needed for the output proof file
        df = df[['Job Code', 'Cost Code', 'Name', 'Hours']]
        
        # create a dataframe that will be the output file 
        csv_df = pd.DataFrame(columns = df.columns)
        # set the first row to be just 'Indirect'
        csv_df.loc[0,'Job Code'] = 'Indirect'
        # get the unique job code & cost code combinations
        # unique_jcc = pd.unique(df['Job Code'] + "<>" +df['Cost Code'])
        # get the different indirect jobs present 
        indirect_jobs = pd.unique(df['Job Code'])
        # start at 0 and then count up 
        indirect_hours = 0
        # do this loop for the indirect jobs
        for job in indirect_jobs:
            # get just that job's data from the indirect df
            chunk = df[df['Job Code'] == job].copy()
            # sum up the hours of that job
            sum_of_chunk = chunk.sum()['Hours']
            # add the sum of the chunk to the indirect horus
            indirect_hours += sum_of_chunk
            # append the df of that job
            csv_df = csv_df.append(chunk, ignore_index=True)
            # append the sum of that job to the bottom of the output df
            csv_df = csv_df.append({'Job Code':'Sum of ' + job, 'Hours':sum_of_chunk}, ignore_index=True)
            # append a blank row after each different indirect job
            csv_df = csv_df.append(pd.Series(dtype=int), ignore_index=True)
        
        # append the total hours of indirect
        csv_df = csv_df.append({'Job Code':'Sum of Indirect', 'Hours':indirect_hours}, ignore_index=True)
        # append a blank row
        csv_df = csv_df.append(pd.Series(dtype=int), ignore_index=True)
        # append another blank row
        csv_df = csv_df.append(pd.Series(dtype=int), ignore_index=True)
        # append the Direct tag to show the start of direct hours
        csv_df = csv_df.append({'Job Code':'Direct'}, ignore_index=True)
        # get the direct horus from the states dict
        df = states_dict[state]['Direct'].copy()
        # only keep the columns needed for output
        df = df[['Job Code', 'Cost Code', 'Name', 'Hours']]
        # get the unique jobs in the direct df
        direct_jobs = pd.unique(df['Job Code'])
        # initialize the total direct hours
        direct_hours = 0
        # go thru each unique direct job
        for job in direct_jobs:
            # get only that direct jobs' data
            chunk = df[df['Job Code'] == job].copy()
            # take the sum of the chunk
            sum_of_chunk = chunk.sum()['Hours']
            # add the sum to the direct hours counter
            direct_hours += sum_of_chunk
            # add the chunk to the data frame
            csv_df = csv_df.append(chunk, ignore_index=True)
            # add the sum of the chunk to the row below 
            csv_df = csv_df.append({'Job Code':'Sum of ' + job, 'Hours':sum_of_chunk}, ignore_index=True)
            # append a blank row beneath each job
            csv_df = csv_df.append(pd.Series(dtype=int), ignore_index=True)
        
        # append the sum of the direct horus at the bottom
        csv_df = csv_df.append({'Job Code':'Sum of Direct', 'Hours':direct_hours}, ignore_index=True)
        # add a blank row
        csv_df = csv_df.append(pd.Series(dtype=int), ignore_index=True)
        # put the states output csv df into the dict
        states_dict[state]['CSV'] = csv_df
        # get that states fablisting entries
        fab_listing = grab_google_sheet(sheet, start_date, end_date)
        # calculate the tonnage
        final_tons = fab_listing.sum()['Weight'] / 2000    
        
        # Enter the states direct hours
        summary.loc[state, 'Direct'] = direct_hours
        # enter the states indirect hours
        summary.loc[state, 'Indirect'] = indirect_hours
        # enter the states total hours
        summary.loc[state, 'Total'] = direct_hours + indirect_hours
        # enter the states amount of OT
        summary.loc[state, 'OT'] = ot[state]
        # enter the states tonnage
        summary.loc[state, 'Tons'] = final_tons
        
        # create the timestamp for the file name 
        csv_file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%m-%S")
        # generate the file name for the output
        csv_detail_summary_name = directory_for_dumping + state + "_" + start_date.replace("/","-") + "_jobs_breakdown_" + csv_file_timestamp + ".csv"
        # create a csv file with this_df
        csv_df.to_csv(csv_detail_summary_name, index=False)    
        
        
        
    server_destination = "X:\\PRODUCTION REPORT FROM US\\Productivity Reports\\Prod Reports 2021\\"
    local_destination = directory_for_dumping.replace('//','\\')
    
    
    # printing shit out to get stuff from the console 
    print('\n')
    print("Go to '2021 - Weekly Shop Hours Report.xlsx' located here:")
    print("\t" + server_destination)   
    subprocess.Popen('explorer ' + server_destination)
    print('And update the numbers from the table below\n\n')
    print(summary)
    print('\n')
    print("Then get the files labeled with 'STATE_"+start_date.replace("/","-")+"_jobs_breakdown':")
    print("\t" + local_destination)  
    subprocess.Popen('explorer ' + local_destination)
    
    print('And attach them to the email that gets sent to the group:')
    print('\t"Weekly Shop Hours Report - Preliminary"') 
    print('\n*Make sure to manually verify the CSM Misc Metal guys*')
    print('\tRobert Burrell, Rodger Grant, & Robert Richardson')
     
    # summary.to_csv(directory_for_dumping + 'Total_summary_' + start_date.replace("/","-") + "_" + csv_file_timestamp + ".csv")
    
    de_list = summary.loc['DE']
    de_list['Date'] = start_date
    summary_de = summary_de.append(de_list, ignore_index=True)
    
    md_list = summary.loc['MD']
    md_list['Date'] = start_date
    summary_md = summary_md.append(md_list, ignore_index=True)

    tn_list = summary.loc['TN']
    tn_list['Date'] = start_date
    summary_tn = summary_tn.append(tn_list, ignore_index=True)    
    
    
    
    start_dt = start_dt + datetime.timedelta(days=7)
    start_date = datetime.datetime.strftime(start_dt, "%m/%d/%Y")    
    

# summary_de.to_csv(directory_for_dumping + 'DE_summary_' + csv_file_timestamp + '.csv')
# summary_md.to_csv(directory_for_dumping + 'MD_summary_' + csv_file_timestamp + '.csv')
# summary_tn.to_csv(directory_for_dumping + 'TN_summary_' + csv_file_timestamp + '.csv')






while True:
    x = input('Type "QUIT" to exit the prompt.\t')
    if x == 'QUIT':
        quit()









