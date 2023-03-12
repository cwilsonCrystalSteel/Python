# -*- coding: utf-8 -*-
"""
Created on Mon Apr 26 12:59:10 2021

@author: CWilson
"""


import glob
import os
import pandas as pd
from bs4 import BeautifulSoup
import datetime
import numpy as np



def new_and_imporved_group_hours_html_reader(html_file, in_and_out_times=False):
    
    # I Dont know what I did but this gets all the rows of the html table
    data = []
    with open(html_file, 'r') as f:
        
        soup = BeautifulSoup(f, 'html.parser')
        # print(soup.prettify())
        table = soup.find('body')
        rows = table.find_all('tr')
    
    
    
        for i,row in enumerate(rows):
            # i+= 1
            # row = rows[i]
            row
            cols = row.find_all('td')
            print(len(cols))
            
            
            # cell_contains_table = [j for j,ele in enumerate(cols) if 'table' in str(ele)]
            # if len(cell_contains_table):
            #     print('removing a table record {}'.format(i))
            #     cols.pop(cell_contains_table[0])
            
            
            if len(cols) == 18 and i == 0:
                print('found header row: {}'.format(i))
                # this is the headers row
                headers = [ele.text.strip() for ele in cols]
                headers.remove('')
                headers.remove('M')
                headers.remove('I')
                headers.remove('O')
                headers.remove('Note')
                headers.remove('Edit')
                headers.remove('Brk')
                headers = ['Name'] + headers
                data.append(headers)
                
            # the colspan tag only appears when there is an employee name i think
            elif any([ele.get('colspan') for ele in cols]) and len(cols) == 1:
                id_name = [ele.text.strip() for ele in cols][0]
                print('This is the employee name row for {} on {}'.format(id_name,i))
            # the only other tie there is only one 'td' is when there is a blank row
            elif len(cols) == 1:
                print('this is an empty row: {}'.format(i))
            else:
                print('this is a record row: {}'.format(i))
                col_texts = [ele.text.strip() for ele in cols]
                
                first_non_blank = next(sub for sub in col_texts if sub)
                
                if first_non_blank == 'X':
                    first_non_blank = next(sub for sub in col_texts[col_texts.index('X')+1:] if sub)
                    print('{} the first nonblank was x, the next was {}'.format(i,first_non_blank))
                
                first_non_blank_idx = col_texts.index(first_non_blank)
                next_data = [id_name]
                for j in range(first_non_blank_idx, len(col_texts)):
                    print(col_texts[j])
                    next_data.append(col_texts[j])
                # timein = col_texts[first_non_blank_idx]
                # actualtimein = col_texts[11]
                # timeout = col_texts[12]
                # actualtimeout = col_texts[13]
                # pto = col_texts[14]
                # jobcode = col_texts[15]
                # costcode = col_texts[16]
                # hours = col_texts[17]
                # rate = col_texts[18]
                # shifttotal = col_texts[19]
                # weektotal = col_texts[20]
                # next_data = [id_name, timein, actualtimein, timeout, 
                #              actualtimeout, pto, jobcode, costcode, hours, 
                #              rate, shifttotal, weektotal]
                data.append(next_data)
                
                print(data)
            
            
            
            
            # classes = [ele.get('class') for ele in cols]
            
            
            
            
            
            
            
            
            # cols = [ele.text.strip() for ele in cols]
            # cols = [ele for ele in cols if ele]
            # # if len(cols) > 2:
            # data.append(cols)
    df = pd.DataFrame(data[1:], columns=data[0])
    df['Cost Code'] = df['Cost Code'].replace('', 'no cost code')
    df['Hours'] = pd.to_numeric(df['Hours'])
    df['Job #'] = pd.to_numeric(df['Job Code'].str.split('-').str[0])
    df['Name'] = df['Name'].str.split(' - ').str[1]
    
    
    lot_ccs = df[df['Cost Code'].str[0] == '9']['Cost Code']
    # get the lot cost codes ending in PAINT or LOAD
    lot_ccs = lot_ccs[(lot_ccs.str[-5:] == 'PAINT') | (lot_ccs.str[-4:] == 'LOAD')]
    # chop off the PAINT or LOAD & strip whitespace
    new_lot_ccs = lot_ccs.str[:-5].str.strip()    
    df.loc[new_lot_ccs.index, 'Cost Code'] = new_lot_ccs
    df = df.rename(columns = {'Time in':'Time In', 'Time out':'Time Out'})
    df = df[['Cost Code', 'Hours', 'Job #', 'Job Code', 'Name', 'Time In', 'Time Out']]
    
    if not in_and_out_times:
        df = df.drop(columns = ['Time In','Time Out'], axis=0)
    
    return df


def output_dict_of_each_employees_hours(html_file):
    
    # I Dont know what I did but this gets all the rows of the html table
    data = []
    with open(html_file, 'r') as f:
        
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
    
    
    # os.remove(html_file)
    
    # fix the header list 
    data[0] = data[0][6:]
    
    new_data = []
    # chops out any empty lists  
    for row in data[1:]:        
        if len(row) != 0:
            new_data.append(row)
    
    # run this 3 times in case there are multiple 'X'
    for xcount in range(0,4):
        # this gets rid of any rows that have an X as the first element
        for i,row in enumerate(new_data):
            if row[0] == 'X':
                new_data[i] = row[1:]     
    
    # create an empty dict that will store name:list_of_hours
    employees = {}
    # iterate thru each row of the list
    for row in new_data:
        # if there is only one element, it is the ID/NAME
        if len(row) == 1:
            # split the string by a dash
            employee_break = row[0].split('-')
            # id is the first part
            emp_id = employee_break[0]
            # name is the second part
            name = employee_break[1][1:]            
            if len(employee_break) >= 2:
                for partial in employee_break[2:]:
                    name = name + '-' + partial
            # create an empty list for that employee key
            employees[name] = []
            # skip to the next row of new_data
            continue
        # append the hours to the key
        employees[name].append(row)
    return employees
 


   
def output_each_clock_entry_job_and_costcode(html_file):
    # get the employees_dict from the output_dict_of_each_employees_hours function
    employees = output_dict_of_each_employees_hours(html_file)
    # this df will be the relational transformation of the HTML file
    times_df = pd.DataFrame()
    
    # create an empty dict to store the [time-in, time-out, # hours, job #] for each employee
    cleansed_employees = {}
    # iterate thru each employee 
    for name in list(employees.keys()):
        # create an empty list to store data in for this row of data
        cleansed_data = []
        # iterate thru the hours of that employee
        for i,row in enumerate(employees[name]):
            # check to see if there is a missing timestamp somewhere in the data
            if 'Missed' in row:
                print("ERROR ERROR ERROR")
                print("Missing time found for : " + name)
                print("Please fix this - data will not be accurate if left this way")
                
                
            # sees if the 0 & 1 items in the row can be converted to a datetime
            try:
                # start time is the first value -> convert from string to datetime
                start = datetime.datetime.strptime(row[0], "%m/%d/%Y %I:%M %p")
                # convert end time to datetime
                end = datetime.datetime.strptime(row[1], "%m/%d/%Y %I:%M %p")
            except:
                continue
            
            # calculate the difference
            time = end - start
            # try to figure out if the item in the list is a job or not
            for i,val in enumerate(row):
                try:
                    # if the first 4 values are an int then that is the job number
                    int(val[:4])
                    job = val
                    job_idx = i
                    job_num = int(job[:4])
                    break
                except:
                    continue
            

            try:
                float(row[job_idx+1])
                cost_code = 'no cost code'
            except:

                cost_code = row[job_idx+1]
            
            
            # determine if there was a break or not on that clock in
            for val in row:
                try:
                    # only proceed if you can strip the 'u' from it and the remaining cna be an int
                    break_length_1 = val.strip('u')
                    break_length = int(break_length_1)
                    # if the first 2 are numbers and the last one is a 'u' and it is 3 long
                    # if val[1].isnumeric() and val[-1] == 'u':
                    #     break_length = val.strip('u')
                        #subtract the 30 minute break from that seciton of work
                    time = time - datetime.timedelta(minutes=break_length)
                    break_length = 0
                except:
                    pass
            
                
            # calculate hours down here b/c 30 minute breaks
            hours = time.total_seconds() / 3600
            # round to 2 decimals to clean up the output
            hours = np.round(hours,2)
            append_list = [name, start, end, hours, job, job_num, cost_code]

            # append to the list
            cleansed_data.append(append_list)
        
        # append all the times/hours/jobs to the employee key 
        cleansed_employees[name] = cleansed_data
        times_df = times_df.append(cleansed_data)
    
    # rename the columns to something meaningful
    times_df = times_df.rename(columns={0:'Name',1:'Start',2:'End',3:'Hours',4:'Job', 5:'Job #',6:'Cost Code'})  
    
    return times_df



def new_output_each_clock_entry_job_and_costcode(html_file, in_and_out_times=False):
    ''' This updated method is simpler and should automatically account for breaks.
        It does not include the start & stop times for each clock tho.
        It returns a dataframe with each unique clock from the html file that is
            from timeclock group hours download.
            
        This relies on the job code being the first item in the row that has 
            characters 0,1,2 being numbers and a hyphen being present in the same item.
    '''
    
    
    # get the employees_dict from the output_dict_of_each_employees_hours function
    employees = output_dict_of_each_employees_hours(html_file)
    # this df will be the relational transformation of the HTML file
    times_df = pd.DataFrame()

    
    
    # loop thru each employee in the employees dict
    for employee in employees.keys():
        # loop thru each list for that employee - each list is that unique clock entry
        for clock in employees[employee]:
            # find the job code - job code has a dash and first 3 are numbers
            job_code = next(i for i in clock if ('-' in i and i[:3].isnumeric()))
            # get the next item in the list
            next_one = clock[clock.index(job_code) + 1]
            # check if the next item is a float -> if it is, then it is the hours
            try:
                float(next_one)
                # set the cost code to be a filler value
                cost_code = 'no cost code'
                # the hours are the item after the job code then
                hours = float(next_one)
            # if the item was not a float, then it was the cost code
            except:
                # cost code is the value after job code then
                cost_code = next_one
                # hours are the next item after cost code
                hours = float(clock[clock.index(cost_code) + 1])
            # create the dict that becomes the newest row of the df 
            append_dict = {'Name' : employee,
                           'Job #': int(job_code.split('-')[0]),
                           'Job Code' : job_code,
                           'Cost Code' : cost_code,
                           'Hours' : hours}
            if in_and_out_times == True:
                append_dict['Time In'] = clock[0]
                append_dict['Time Out'] = clock[1]
            # append the new row to the df
            times_df = times_df.append(append_dict, ignore_index=True)
    

    ''' this tidbit removes the 'PAINT' or 'LOAD' portion of LOT COST CODES '''
    # get the lot cost codes (always start with a 9)
    lot_ccs = times_df[times_df['Cost Code'].str[0] == '9']['Cost Code']
    # get the lot cost codes ending in PAINT or LOAD
    lot_ccs = lot_ccs[(lot_ccs.str[-5:] == 'PAINT') | (lot_ccs.str[-4:] == 'LOAD')]
    # chop off the PAINT or LOAD & strip whitespace
    new_lot_ccs = lot_ccs.str[:-5].str.strip()
    # set the new lot cost codes back into the dataframes
    times_df.loc[new_lot_ccs.index, 'Cost Code'] = new_lot_ccs
    
    
    # return the dataframe
    return times_df



def output_group_hours_by_job_code(html_file):
    
    times_df = output_each_clock_entry_job_and_costcode(html_file)
    

    # get the different jobs in the dataframe
    jobs = pd.unique(times_df['Job'])
    # get the different names in the dataframe
    names = pd.unique(times_df['Name'])
    
    
    # initialize a dataframe starting with just the names 
    output_df = pd.DataFrame(columns=['Name'], data=names, index=names)
    # initialize empty columns for each job code
    for job in jobs:
        output_df[str(job) + ' Hours'] = 0
    
    # iterate thru each name 
    for name in names:
        # get the portion of the dataframe for that name
        big_chunk = times_df[times_df['Name'] == name]
        # get their unique jobs worked
        their_jobs = pd.unique(big_chunk['Job'])
        # iterate thru each of the unique jobs they worked
        for job in their_jobs:
            # break their df into a smaller job-specific df
            chunk = big_chunk[big_chunk['Job'] == job]
            # sum the hours worked for that job
            hours_for_job = chunk.sum()['Hours']
            # round the hours to 2 decimal places
            hours_for_job = np.round(hours_for_job, 2)
            # get the column name based on the job
            col_name = str(job) + ' Hours'
            # place their hours in the rightful spot in the output dataframe
            output_df.loc[name,col_name] = hours_for_job
        
        
    
    return output_df





















