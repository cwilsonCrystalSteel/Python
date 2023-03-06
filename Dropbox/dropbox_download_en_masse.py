# -*- coding: utf-8 -*-
"""
Created on Fri Nov 12 15:37:58 2021

@author: CWilson
"""
import os 
import glob
import pandas as pd
from navigate_EVA_folder_function import get_df_of_all_lots_files_information

''' This is to DOWNLOAD ALL LOTS en mass '''
# turn it to false to prevent any accidents
download_and_append_en_masse = True



df = get_df_of_all_lots_files_information()
# a = df[df['shop'] == 'CSM']
# b = a[a['job'] == 2038]
# c = b[b['lot'] == 'T093']
# d0 = pd.read_excel(c['destination'].iloc[0], header=2, engine='xlrd', sheet_name='RAW DATA')
# d1 = pd.read_excel(c['destination'].iloc[0], header=2, engine='xlrd', sheet_name='RAW DATA')
# x = d1[d1['PAGE'] == 'AN8021']

# base_dir = "C:\\Users\\cwilson\\Dropbox\\EVA REPORTS FOR THE DAY\\"
base_dir = 'X:\\production control\\EVA REPORTS FOR THE DAY\\'






try:
    years = os.listdir(base_dir)
except:
    print('COULD NOT CONNECT TO THE X DRIVE / DROPBOX')
    exit()
    
    
    
    
if 'desktop.ini' in years:
    years.remove('desktop.ini')

critical_columns = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'PRODUCTION CODE', 'QTY','SHAPE', 'LABOR CODE', 'MAIN MEMBER', 'TOTAL MANHOURS']


jobs = {}

for year in years:
    
    months = os.listdir(base_dir + year)
    
    for month in months:
        
        days = os.listdir(base_dir + year + '\\' + month)
        
        for day in days:
            
            this_day = os.listdir(base_dir + year + '\\' + month + '\\' + day)
            
            xls_files = glob.glob(base_dir + year + '\\' + month + '\\' + day + "\\*.xls")
            
            for xls_file in xls_files:
                
                basename = os.path.basename(xls_file)
                
                basename_components = basename.split('-')
                
                if len(basename_components) == 3:
                    # job # is the start of the filename
                    job_str = basename_components[0]
                    #
                    job_int = int(job_str)
                    # lot is the middle portion of the filename
                    lot = basename_components[1]
                    # chop off the '.xls' and only maintain the shop 
                    shop = basename_components[2][:-4]
                    
                    # skip the xls file if it does not have a valid name
                    if lot[0] != 'T':
                        continue
                
                else:
                    print('ERROR THE FILENAME IS INVALID')
                    ''' SEND ERROR THAT THE FILENAME IS INVALID
                    1) Emad
                    2) mildred tong
                    3) me
                    '''
                print(day, month, year)
                
                
                
                # open the current lots file -> fingers crossed the header is alwasy row 2
                # only maintain those 9 columns that we actually need to pass along -> smaller file sizes
                try:
                    xls_lot = pd.read_excel(xls_file, header=2, engine='xlrd', sheet_name='RAW DATA', usecols=critical_columns)
                except:
                    print('THERE WAS AN ERROR OPENING THE XLS FILE IN THE DROPBOX FOLDER')
                    print(day, month, year, basename)
                    # input('press ENTER to continue')
                    continue
                xls_main_name = 'c://downloads//' + job_str + '.xlsx'
                
                
                
               # this will be how the cron job operation runs I think
                if not download_and_append_en_masse:
                     #if the file for that job exists, open it
                    if os.path.exists(xls_main_name):
                        
                        # start by assuming headers are on row 0
                        header_num = 0
                        # open just the first col with 5 rows
                        xls_main_test = pd.read_excel(xls_main_name, engine='xlrd', header=header_num, nrows=5, usecols=[0])
                        # if the header is not 'JOB NUMBER' iterate header_num and try again
                        while xls_main_test.columns[0] != 'JOB NUMBER':
                            header_num += 1
                            xls_main_test = pd.read_excel(xls_main_name, engine='xlrd', header=header_num, nrows=5, usecols=[0])
    
                        #open the main file with the correct header number
                        xls_main = pd.read_excel(xls_main_name, engine='xlrd', header=header_num)
                        # append the lot's pieces to that main file 
                        xls_main = xls_main.append(xls_lot)
                            
                        xls_main.to_excel(xls_main_name, index=False)
                        
                    
                    #if the file for that job DOES NOT EXIST, create it & create the xls_main variable as that LOTS data
                    else:
                        
                        xls_lot.to_excel(xls_main_name, index=False)
                        
                        xls_main = xls_lot.copy()
              
                
              
                
              
                if download_and_append_en_masse:
                        
                    if xls_main_name in jobs.keys():
                        df = jobs[xls_main_name]
                        
                        df = df.append(xls_lot)
                        
                        jobs[xls_main_name] = df      
                    else:
                        jobs[xls_main_name] = xls_lot


# this now puts the dfs into .xls files in the c:\downloads folder from the dict 
if download_and_append_en_masse:
    print('Now creating all the .xlsx files...')
    for file in jobs.keys():
    
        df = jobs[file]
        
        print(file)
        
        df.to_excel(file, index=False)
                