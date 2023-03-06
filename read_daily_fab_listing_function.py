# -*- coding: utf-8 -*-
"""
Created on Mon Dec  7 08:15:24 2020

@author: CWilson
"""

''' 
Let this file be the function for reading the indivual sheets of the 
daily fab listing excel files

This function will do the following:
    1) break the sheet into 2 dataframes, one for each column group
    2) check if the sheet has welders present or not
    3) obtain the job & seq number 
    4) for each job/seq, it will grab the QTY, PC MARK, total WT, fitter #, welder #
    5) return a dataframe that has the following columns for every item:
        job #, seq/lot, qty, pcmark, wt, fitter, welder
        
The inputs needed for this function will be the excel file
'''


import pandas as pd
from datetime import datetime
from os import listdir
from os.path import isfile, join
import numpy as np


def yield_todays_output(df_in, date):
# Function to perform the compiling of data from df_left and df_right
# call the func so that: this_func(df_left) and this_func(df_right)
# append the data to the output dataframe

    func_output = pd.DataFrame(columns=['Date', 'Job #', 'Seq #', 'Pc Mark', 'Qty', 'Wt', 'Fitter', 'QCF', 'Welder', 'QCW'])
    
    for idx in df_in.index:
        
        # Retains the job#/seq# from the previous idx
        if pd.isnull(df_in.loc[idx, 'Job #']):
            pass
        elif str(df_in.loc[idx, 'Job #']).strip() == '':
            pass
        else:
            this_job_seq = df_in.loc[idx, 'Job #']
        
        # Create values for the job# and seq# based on the column 'Job #'
        if '-' in str(this_job_seq):
            job_num = str(this_job_seq).split('-')[0]
            seq_num = str(this_job_seq).split('-')[1]
            try:
                job_num = int(job_num)
            except:
                pass
            try:
                seq_num = int(seq_num)
            except:
                pass
        else:
            job_num = this_job_seq
            seq_num = None
            # job_num = int(job_num)
        
        # Assigning values from the dataframes
        qty = df_in.loc[idx, 'Qty']
        pcmark = df_in.loc[idx, 'Pc Mark']
        wt = df_in.loc[idx, 'Weight']
        fitter = df_in.loc[idx, 'Fitter']
        qcf = df_in.loc[idx, 'QCF']
        
        if 'Welder' in df_in.columns:
            welder = df_in.loc[idx, 'Welder']
            qcw = df_in.loc[idx, 'QCW']
        else:
            welder = None
            qcw = None
    
        # Appending row to the output dataframe for this piece
        
        this_pc = {'Date':date,
                   'Job #':job_num, 
                   'Seq #':seq_num, 
                   'Pc Mark':pcmark, 
                   'Qty':qty, 
                   'Wt':wt, 
                   'Fitter':fitter, 
                   'QCF':qcf,
                   'Welder':welder,
                   'QCW':qcw}
        
        func_output = func_output.append(this_pc, ignore_index=True)
    
    return func_output


def compile_daily_fab(folder, file):
    
    
    output = pd.DataFrame(columns=['Date', 'Job #', 'Seq #', 'Pc Mark', 'Qty', 'Wt', 'Fitter', 'QCF', 'Welder', 'QCW'])

    # Obtains the year and month from the file name
    year = file[:4]
    month = file[5:7]

    
    
    # Opens the excel file
    xl = pd.ExcelFile(folder + '\\' + file)
    
    # creates a list of the names of all the sheets in the xcel file
    sheets = xl.sheet_names
    
    # day = sheets[6]
    for day in sheets:
        
        # only reads the sheet if it is type(int)
        try:
            int(day)
        except:
            continue
        
        df = xl.parse(str(day))
        
        
        if len(day) == 1:
            day = '0' + day
            
        try:
            date = datetime.strptime(year + '-' + month + '-' + day, '%Y-%m-%d')
        except:
            continue
        
        # Identifying the row index where the data begins
        column_idx = df.loc[df[df.columns[0]] == 'Job #'].index[0]
        cols = df.loc[column_idx].fillna('split here').tolist()
        
        
        # Seperating the columns
        split_here_idx = cols.index('split here')
        df_left_cols = cols[:split_here_idx]
        df_right_cols = cols[split_here_idx+1:]
        
        
        # Renaming the 'QC' columns to their respective fit or weld category
        df_left_cols[df_left_cols.index('Fitter') + 1] = 'QCF'
        df_right_cols[df_right_cols.index('Fitter') + 1] = 'QCF'
        if 'Welder' in df_left_cols:
            df_left_cols[df_left_cols.index('Welder') + 1] = 'QCW'
        if 'Welder' in df_right_cols:
            df_right_cols[df_right_cols.index('Welder') + 1] = 'QCW'
        
        # Appending a ' 1' to the end of the right columns -> unique column names
        df_right_cols = [str(x) + ' 1' for x in df_right_cols]
        
        # The new columns 
        cols = df_left_cols + ['split here'] + df_right_cols
        
        # Chooping off unused initial rows
        df = df.loc[column_idx + 1:].reset_index()
        df = df.drop([df.columns[0]], axis=1)
        
        # Changing the names of columns in df to useful names
        df.columns = cols
        
        # In case of extra blank columns that got named 'split here'
        df_right_cols = [i for i in df_right_cols if i[:5] != 'split']
        
        # Splitting the sheet into the left and right columns
        df_left = df[df_left_cols]
        df_right = df[df_right_cols]
        
        # Making the column names the same between both left and right df's
        df_right.columns = df_left.columns
        
        
        # Chopping off rows that do not have a QCF value
        df_left = df_left[df_left['QCF'].notna()]
        df_right = df_right[df_right['QCF'].notna()]
        
        
        
        
        # Appends the left and right column data to the output dataframe
        output = output.append(yield_todays_output(df_left, date), ignore_index=True)
        output = output.append(yield_todays_output(df_right, date), ignore_index=True)

    xl.close()
    return output









# from os import listdir
# from os.path import isfile, join
# daily_fab_python_folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
# files = [f for f in listdir(daily_fab_python_folder) 
#             if isfile(join(daily_fab_python_folder,f))]


# folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
# file = "2020.09 Daily Fabrication.xlsx"






    
    



# out = compile_daily_fab(folder, file)
























