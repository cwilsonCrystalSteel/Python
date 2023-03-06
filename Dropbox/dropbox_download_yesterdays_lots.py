# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 09:10:02 2022

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
import os
import glob
import pandas as pd
import datetime

log_path = 'C:\\Users\\cwilson\\Documents\\Python\\Dropbox\\Last_day_retrieved_log.csv'
def write_to_logfile(dt, status):
    dt_log_string = dt.strftime('%Y-%m-%d')
    log = pd.read_csv(log_path)
    log = log.append({'date':dt_log_string, 'status':status}, ignore_index=True)
    log.to_csv(log_path, index=False)


critical_columns = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'PRODUCTION CODE', 'QTY','SHAPE', 'LABOR CODE', 'MAIN MEMBER', 'TOTAL MANHOURS']
# base_dir = "C:\\Users\\cwilson\\Dropbox\\EVA REPORTS FOR THE DAY\\"
base_dir = 'X:\\production control\\EVA REPORTS FOR THE DAY\\'



# open up the log 
log = pd.read_csv(log_path)
# get the last time a file was the status
last_successful_date = log[log['status'].str.startswith('X')].iloc[-1]['date']
# get the datetime of that date
last_successful_dt = datetime.datetime.strptime(last_successful_date,'%Y-%m-%d').date()
# get yesterday's date
yesterday = (datetime.datetime.now() + datetime.timedelta(days=-1)).date()
# get the number of days between last success and yesterday
delta = (yesterday - last_successful_dt).days
 

write_to_logfile(yesterday, 'Started')

try:
    years = os.listdir(base_dir)
except:
    print('COULD NOT CONNECT TO THE X DRIVE / DROPBOX')
    write_to_logfile(yesterday, 'Could not connect to the dropbox')
    exit()
    

year = str(yesterday.year)

months = os.listdir(base_dir + year)

#get it as a zero padded number -> 01 or 12
month_num = str(yesterday.month).zfill(2)
# '%B' gets the month name
month_name = yesterday.strftime('%B').upper()

month_str = month_num + ' - ' + month_name

days = os.listdir(base_dir + year + '\\' + month_str)





for i in range(0,delta):
    day_dt = last_successful_dt + datetime.timedelta(days=1+i)
    
    
    # get yesterday as the MMDDYYYY format that dropbox has
    day = day_dt.strftime('%m%d%Y')
    print(day)
    
    
    
    
    
    if day in days:
        
        xls_files = glob.glob(base_dir + year + '\\' + month_str + '\\' + day + "\\*.xls")
        
        for xls_file in xls_files:
            
            basename = os.path.basename(xls_file)
            
            basename_components = basename.split('-')
            
            if len(basename_components) == 3:
                # job # is the start of the filename
                job_str = basename_components[0]
                # lot is the middle portion of the filename
                lot = basename_components[1]
                # chop off the '.xls' and only maintain the shop 
                shop = basename_components[2][:-4]
                
                
            else:
                print('ERROR THE FILENAME IS INVALID')
                ''' SEND ERROR THAT THE FILENAME IS INVALID
                1) Emad
                2) mildred tong
                3) me
                '''
            print(basename, day, month_str, year)
            
            # open the current lots file -> fingers crossed the header is alwasy row 2
            # only maintain those 9 columns that we actually need to pass along -> smaller file sizes
            try:
                xls_lot = pd.read_excel(xls_file, header=2, engine='xlrd', sheet_name='RAW DATA', usecols=critical_columns)
            except:
                print('the file does not have a valid sheet_name to be added')
                print('\t\t',basename, day, month_str, year)
                continue
            xls_main_name = 'c://downloads//' + job_str + '.xlsx'  
    
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
                    
                xls_main.to_excel(xls_main_name , index=False)
                
                print('Successfully added {0} to {1}'.format(basename, xls_main_name))
                
            
            #if the file for that job DOES NOT EXIST, create it & create the xls_main variable as that LOTS data
            else:
                
                xls_lot.to_excel(xls_main_name, index=False)
                
                xls_main = xls_lot.copy()
            
            log_file = open("C:\\Users\\cwilson\\Documents\\Python\\Dropbox\\Log of LOTS.txt","a")
            log_file.write(xls_file + ", " + datetime.datetime.now().strftime('%m/%d/%Y %H:%M') + " \n")
            log_file.close()
            
            write_to_logfile(yesterday, xls_file)

