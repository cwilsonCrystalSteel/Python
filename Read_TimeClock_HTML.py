# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 09:34:57 2021

@author: CWilson
"""

import glob
import os
import pandas as pd


def return_direct_indirect_sums_dfs(state='TN', folder="C://Users//Cwilson//downloads//"):

    # Grab all HTML files in downloads
    list_of_files = glob.glob(folder + "*.html") # * means all if need specific format then *.csv
    # Create a list with only the states we want to look at
    state_files = [f for f in list_of_files if state + " Job Code Summary" in f]
    # Get the most recent file for that state
    latest_file = max(state_files, key=os.path.getctime)
    print(latest_file)
    
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
    na_index = df[df['Job Code'].isnull()].index.tolist()
    # Seperate the indirect by the top half of the table. Remove nan rows
    indirect = df[:na_index[0]-2]
    # Seperate direct section as the bottom half. Remove nan rows
    direct = df[na_index[0]:-2]
    
    
    # Special tasks for each facility
    if state == 'TN':
        print('TN special tasks: \n\tMove 302xxx from direct to indirect\n\tDrop 1824')
        # Move 302xxx to the indirect df and remove from direct df
        # Completely remove 1824 from the working hours
        for i in direct.index:
            job_num = direct['Job Code'][i][:4]
            if (job_num == '3020') or (job_num == '3021'):
                indirect = indirect.append(direct.loc[i])
                direct = direct.drop([i])
                
            if job_num == '1824':
                direct = direct.drop([i])
                
    elif state == 'MD':
        print('MD Special Tasks: Remove 1626 from direct')
        for i in direct.index:
            job_num = direct['Job Code'][i][:4]
            if job_num == '1626':
                indirect = indirect.append(direct.loc[i])
                direct = direct.drop([i])
        
    elif state == 'DE':
        print("DE special tasks")
    
        
    for column in direct.columns[1:].tolist():
        direct.loc[:, column] = direct.loc[:, column].astype(float).copy()
        indirect.loc[:, column] = indirect.loc[:, column].astype(float).copy()
    
    
    direct_sum = direct[direct.columns[1:]].sum(axis=0)
    indirect_sum = indirect[indirect.columns[1:]].sum(axis=0)
    
    sums = pd.DataFrame(data=[direct_sum, indirect_sum], index=['Direct', 'Indirect'])

    
    return [direct, indirect, sums]