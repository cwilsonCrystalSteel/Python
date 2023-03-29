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
import re

log_path = 'C:\\Users\\cwilson\\Documents\\Python\\Dropbox\\Last_day_retrieved_log.csv'
def write_to_logfile(dt, status):
    dt_log_string = dt.strftime('%Y-%m-%d')
    log = pd.read_csv(log_path)
    log = log.append({'date':dt_log_string, 'status':status}, ignore_index=True)
    log.to_csv(log_path, index=False)


critical_columns = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'PRODUCTION CODE', 'QTY','SHAPE', 'LABOR CODE', 'MAIN MEMBER', 'TOTAL MANHOURS', 'WEIGHT']
# base_dir = "C:\\Users\\cwilson\\Dropbox\\EVA REPORTS FOR THE DAY\\"
base_dir = 'X:\\production control\\EVA REPORTS FOR THE DAY\\'
base_dir = '\\\\192.168.50.9\Dropbox_(CSF)\\production control\\EVA REPORTS FOR THE DAY\\'


# open up the log 
log = pd.read_csv(log_path)
try:
    # get the last time a file was the status
    last_successful_date = log[log['status'].str.startswith('X')].iloc[-1]['date']
    last_successful_date = log[log['status'] == 'Successfully completed all files'].iloc[-1]['date']
except Exception:
    last_successful_date = '2021-01-13' # this is the day before the EVA hours started showing up in X -drive
    
# get the datetime of that date
try:
    last_successful_dt = datetime.datetime.strptime(last_successful_date,'%m/%d/%Y').date()
except Exception:
    last_successful_dt = datetime.datetime.strptime(last_successful_date,'%Y-%m-%d').date()

# get yesterday's date
yesterday = (datetime.datetime.now() + datetime.timedelta(days=-1)).date()
# get the number of days between last success and yesterday
delta = (yesterday - last_successful_dt).days
 

# write_to_logfile(yesterday, 'Started')

# try:
#     years = os.listdir(base_dir)
# except:
#     print('COULD NOT CONNECT TO THE X DRIVE / DROPBOX')
#     write_to_logfile(yesterday, 'Could not connect to the dropbox')
#     exit()
    

# year = str(yesterday.year)

# months = os.listdir(base_dir + year)

# #get it as a zero padded number -> 01 or 12
# month_num = str(yesterday.month).zfill(2)
# # '%B' gets the month name
# month_name = yesterday.strftime('%B').upper()

# month_str = month_num + ' - ' + month_name

# days = os.listdir(base_dir + year + '\\' + month_str)





for i in range(0,delta):
    day_dt = last_successful_dt + datetime.timedelta(days=1+i)
    
    day = day_dt.strftime('%m%d%Y')
    print(day)


    write_to_logfile(day_dt, 'Started')
    
    try:
        base_dir = 'X:\\production control\\EVA REPORTS FOR THE DAY\\'
        years = os.listdir(base_dir)
    except:
        print('COULD NOT CONNECT TO THE X DRIVE / DROPBOX')
        try:
            base_dir = '\\\\192.168.50.9\Dropbox_(CSF)\\production control\\EVA REPORTS FOR THE DAY\\'
            years = os.listdir(base_dir)   
        except:
            write_to_logfile(day_dt, 'Could not connect to the dropbox')
            exit()    
    
    year = str(day_dt.year)

    months = os.listdir(base_dir + year)
    
    #get it as a zero padded number -> 01 or 12
    month_num = str(day_dt.month).zfill(2)
    # '%B' gets the month name
    month_name = day_dt.strftime('%B').upper()
    
    month_str = month_num + ' - ' + month_name
    
    days = os.listdir(base_dir + year + '\\' + month_str)
    
    # reread the log in each loop
    log = pd.read_csv(log_path)
    
    already_retrieved_files = pd.unique(log['status'])
    
    if day in days:
    
        xls_files = glob.glob(base_dir + year + '\\' + month_str + '\\' + day + "\\*.xls")
        
        for xls_file in xls_files:
            # xls_file = [i for i in xls_files if 'T084' in i][0]
            xls_file_exists_with_previous_date = False
            # check to see if we have already recorded the xls file before
            if xls_file in already_retrieved_files:
                dates_for_file = log[log['status'] == xls_file]
                # set the indicator true 
                xls_file_exists_with_previous_date = True
                # check to see if we have the same date for that file
                if dates_for_file[dates_for_file['date'] == day_dt.strftime('%Y-%m-%d')].shape[0]:
                    # skip this file if it already exists with that date
                    continue
            
            basename = os.path.basename(xls_file)
            
            basename_components = basename.split('-')
            
            if len(basename_components) != 3:
                print('ERROR THE FILENAME IS INVALID: {}'.format(basename))
                ''' SEND ERROR THAT THE FILENAME IS INVALID
                1) Emad
                2) mildred tong
                3) me
                '''
                try:
                    basename_components = re.split('-|\ ', basename)
                    job_str = basename_components[0]
                    lot = basename_components[1]
                    shop = basename_components[2][:-4]
                    print('Able to resolve the invalid filename')
                except:
                    # 5 + '5'
                    continue
                
            else:
                # job # is the start of the filename
                job_str = basename_components[0]
                # lot is the middle portion of the filename
                lot = basename_components[1]
                # chop off the '.xls' and only maintain the shop 
                shop = basename_components[2][:-4]
                
                
            print(basename, day, month_str, year)
            
            # open the current lots file -> fingers crossed the header is alwasy row 2
            # only maintain those 9 columns that we actually need to pass along -> smaller file sizes
            sheet_num = 0
            while sheet_num < 5:
                try:
                    xls_lot = pd.read_excel(xls_file, header=2, engine='xlrd', sheet_name=sheet_num, usecols=critical_columns)
                    break
                except:
                    print('Sheet Number {} did not contain the critical_cols'.format(sheet_num))
                    print('\t\t',basename, day, month_str, year)
                sheet_num += 1
            if sheet_num  == 4:
                print('the file does not have a valid sheet_name to be added')
                print('\t\t',basename, day, month_str, year)
                # 5 + '5'
                continue                
            # try:
            #     xls_lot = pd.read_excel(xls_file, header=2, engine='xlrd', sheet_name='RAW DATA', usecols=critical_columns)
            # except:
            #     print('the file does not have a valid sheet_name to be added')
            #     print('\t\t',basename, day, month_str, year)
            #     5 + '5'
            #     continue
            
            xls_main_name = 'c://downloads//' + job_str + '.xlsx'  
            # this is if the storage file in c:/downloads exists for that job
            if os.path.exists(xls_main_name):
                
                # start by assuming headers are on row 0
                header_num = 0
                # open just the first col with 5 rows
                try:
                    xls_main_test = pd.read_excel(xls_main_name, engine='xlrd', header=header_num, nrows=5, usecols=[0])
                except:
                    xls_main_test = pd.read_excel(xls_main_name, engine='openpyxl', header=header_num, nrows=5, usecols=[0])
             
                # if the header is not 'JOB NUMBER' iterate header_num and try again
                while xls_main_test.columns[0] != 'JOB NUMBER':
                    header_num += 1
                    try:
                        xls_main_test = pd.read_excel(xls_main_name, engine='xlrd', header=header_num, nrows=5, usecols=[0])
                    except:
                        xls_main_test = pd.read_excel(xls_main_name, engine='openpyxl', header=header_num, nrows=5, usecols=[0])
    
                #open the main file with the correct header number
                try:
                    xls_main = pd.read_excel(xls_main_name, engine='xlrd', header=header_num)
                except:
                    xls_main = pd.read_excel(xls_main_name, engine='openpyxl', header=header_num)
                # add the lot to the df
                xls_lot['LOT'] = lot
                
                if xls_main[xls_main['LOT'] == lot].shape[0]:
                    xls_main = xls_main[xls_main['LOT'] != lot]
                
                xls_lot_grouped = xls_lot.groupby(['JOB NUMBER','SEQUENCE','PAGE','MAIN MEMBER','PRODUCTION CODE','SHAPE','LABOR CODE','LOT']).sum()
                xls_lot_grouped = xls_lot_grouped.reset_index(drop=False)
                # this is for when a lot already exists in the xls_main but needs to be updated
                # if xls_file_exists_with_previous_date:
                #     print('Updating lot {}'.format(lot))
                #     # get a copy of the main file's data for that LOT
                #     xls_main_with_same_lot = xls_main[xls_main['LOT'] == lot].copy()
                #     # create a copy of the xls_lot
                #     xls_lot_to_compare = xls_lot.copy()
                #     # give the xls_main an OLD tag
                #     xls_main_with_same_lot['new/old'] = 'old'
                #     #give the xls_lot copy a NEW tag
                #     xls_lot_to_compare['new/old']='new'
                #     # these are the fields we use to differentiate duplicates
                #     subset = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'MAIN MEMBER','PRODUCTION CODE','SHAPE', 'LABOR CODE', 'LOT']
                
                #     # append the copies of main & xls_lot
                #     x_appeneded = xls_lot_to_compare.append(xls_main_with_same_lot, ignore_index=True)
                #     # get any records that are not duplicates
                #     # this will maintain any records that are unique to xls_main and to xls_lot
                #     non_duplicated = x_appeneded[x_appeneded.duplicated(subset) == False]
                #     # get the new version of any duplicate records
                #     duplicated_and_new = x_appeneded[(x_appeneded.duplicated(subset) == True) & (x_appeneded['new/old'] == 'new')]
                #     # create the updated records for that lot
                #     x_updated_lot = non_duplicated.append(duplicated_and_new) 
                #     # get rid of the new/old column
                #     x_updated_lot = x_updated_lot.drop(columns=['new/old'])
                #     # drop all of the previous lot data
                #     xls_main = xls_main[xls_main['LOT'] != lot]
                #     # append the updated lot data to xls_main
                #     xls_main = xls_main.append(x_updated_lot)
                    
                    
                    
                    
                    
                # append the lot's pieces to that main file 
                xls_main = xls_main.append(xls_lot)
                # send back to excel
                xls_main.to_excel(xls_main_name , index=False)
                
                print('Successfully added {0} to {1}'.format(basename, xls_main_name))
                
            
            #if the file for that job DOES NOT EXIST, create it & create the xls_main variable as that LOTS data
            else:
                # add the lot to the df
                xls_lot['LOT'] = lot
                
                xls_lot.to_excel(xls_main_name, index=False)
                
                xls_main = xls_lot.copy()
            
            log_file = open("C:\\Users\\cwilson\\Documents\\Python\\Dropbox\\Log of LOTS.txt","a")
            log_file.write(xls_file + ", " + datetime.datetime.now().strftime('%m/%d/%Y %H:%M') + " \n")
            log_file.close()
            
            write_to_logfile(day_dt, xls_file)
    write_to_logfile(day_dt, 'Successfully completed all files')

