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
        ph = [i for i in xls_files if 'PH' in i]
        if len(ph):
            print(ph)
        xls_files = [i for i in xls_files if not 'PH' in i]
        
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
                    continue
                
            else:
                # job # is the start of the filename
                job_str = basename_components[0]
                # lot is the middle portion of the filename
                lot = basename_components[1]
                # chop off the '.xls' and only maintain the shop 
                shop = basename_components[2][:-4]
                
                
            # print(basename, day, month_str, year)
            
            ''' Try to massage the lot name '''
            if lot[0] == 'T' and len(lot) == 4:
                print('Rule 0: {}'.format(lot))
                
            elif lot[0] == 'T' and len(lot) == 3:
                old_lot = lot
                lot = 'T' + lot[1:].zfill(3)
                print('Rule 1: {} to {}'.format(old_lot, lot))
                
            elif lot.isnumeric():
                old_lot = lot
                lot = 'T' + lot.zfill(3)
                print('Rule 2: {} to {}'.format(old_lot, lot))
                
            elif lot[0] == 'T' and lot[1:4].isnumeric():
                old_lot = lot
                lot = lot[:4]
                print('Rule 3: {} to {}'.format(old_lot, lot))
                
            else:
                print('This lot does not fill other criteria')
                5 + '5'
                
            
            # we are going to iterate thru the sheet_names as numbers to find the right one b/c calling by name results in a lot of failures
            sheet_num = 0
            max_sheet_num = 6
            while sheet_num <= max_sheet_num:
                try:
                    # open the current lots file -> fingers crossed the header is alwasy row 2
                    # only maintain those 9 columns that we actually need to pass along -> smaller file sizes
                    xls_lot = pd.read_excel(xls_file, header=2, engine='xlrd', sheet_name=sheet_num, usecols=critical_columns)
                    # if for some reason it pulls it but the dataframe is shape (0,0)
                    if not xls_lot.size:
                        5 + '5' #???
                        # keep going in the while loop - try another sheet_name
                        continue
                        
                    break
                except:
                    5
                    # print('Sheet Number {} did not contain the critical_cols'.format(sheet_num))
                    # print('\t\t',basename, day, month_str, year)
                
                sheet_num += 1
                
            # when we max out the sheet_name iterator - skip this LOT
            if sheet_num  == max_sheet_num:
                # print('the file does not have a valid sheet_name to be added')
                # print('\t\t',basename, day, month_str, year)
                5 + '5' #???
                continue                

            # get the 'database' from c:/downlaods/jobnumber
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
                
                # if that LOT already exists in xls_main
                if xls_main[xls_main['LOT'] == lot].shape[0]:
                    # get rid of that LOT's records from xls_main
                    xls_main = xls_main[xls_main['LOT'] != lot]
                
                xls_lot_grouped = xls_lot.groupby(['JOB NUMBER','SEQUENCE','PAGE','MAIN MEMBER','PRODUCTION CODE','SHAPE','LABOR CODE','LOT']).sum()
                xls_lot_grouped = xls_lot_grouped.reset_index(drop=False)
                #
                    
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

