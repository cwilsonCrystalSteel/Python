# -*- coding: utf-8 -*-
"""
Created on Fri Apr 23 09:34:59 2021

@author: CWilson
"""

from TimeClock_Individual_Hours import download_employee_hours
import glob
import os
import pandas as pd
from bs4 import BeautifulSoup
import datetime


def get_employee_hours_html(employee_name, start_date, end_date):
    count = 0
    while count <= 5:
        try:
            download_employee_hours(employee_name, start_date, end_date)
            # Grab all HTML files in downloads
            list_of_htmls = glob.glob("C://Users//Cwilson//downloads//*.html") # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            employee_hours_htmls = [f for f in list_of_htmls if "Hours" in f]
            # Get the most recent file for that state
            latest_employee_hours = max(employee_hours_htmls, key=os.path.getctime)
            print('Newest employee department: ', latest_employee_hours)         
            # delete the file from the computer once it is in memory
            # os.remove(latest_employee_hours)
            count += 1
            print(count)
            break
        except:
            pass
    return latest_employee_hours



def return_employee_hours_per_job(html_file, employee_name):

    # html_file = get_employee_hours_html(employee_name, start_date, end_date)
    
    
    
    #%% I Dont know what I did but this gets all the rows of the html table
    
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
    #%%
    
    # os.remove(html_file)
    
    # fix the header list 
    data[0] = data[0][6:]
    #
    new_data = []
    # chops out any empty lists  
    for row in data:        
        if len(row) != 0:
            new_data.append(row)
    
    # run this 3 times in case there are multiple 'X'
    for xcount in range(0,4):
        # this gets rid of any rows that have an X as the first element
        for i,row in enumerate(new_data):
            if row[0] == 'X':
                new_data[i] = row[1:]     
    
    
    
    
    
    
    cleansed_data = []
    for i,row in enumerate(new_data[1:]):
    
        start = datetime.datetime.strptime(row[0], "%m/%d/%Y %I:%M %p")
        end = datetime.datetime.strptime(row[1], "%m/%d/%Y %I:%M %p")
        time = end - start
        
        for val in row:
            try:
                # if the first 4 values are an int then that is the job number
                int(val[:4])
                job = int(val[:4])
                break
            except:
                continue
        
        
        for val in row:
            try:
                # only proceed if the first 2 are numbers
                int(val[:2])
                # if the first 2 are numbers and the last one is a 'u' and it is 3 long
                if val[:2].isnumeric() and val[-1] == 'u' and len(val) == 3:
                    #subtract the 30 minute break from that seciton of work
                    time = time - datetime.timedelta(minutes=30)
            except:
                pass
                
        
        hours = time.total_seconds() / 3600
        
        cleansed_data.append([start, end, hours, job])
        
    
    
    
    time_out_df = pd.DataFrame(columns=['In','Out','Hours', 'Job'], data = cleansed_data)
    
    jobs = pd.unique(time_out_df['Job'])
    
    output_df = pd.DataFrame(columns=['Name'], data=[employee_name])
    
    for job in jobs:
    
        chunk = time_out_df[time_out_df['Job'] == job]
        hours_for_job = chunk.sum()['Hours']
        col_name = str(job) + ' Hours'
        output_df.loc[0,col_name] = hours_for_job
    
    
    
    return output_df



































