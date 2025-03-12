# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 09:10:02 2022

@author: CWilson
"""
import os
from pathlib import Path
import glob
import pandas as pd
import datetime
import re


# all_LOTs_imported_path = Path(os.getcwd()) / 'Dropbox' / 'Log of LOTS.txt'
log_path = Path(os.getcwd()) / 'Dropbox' / 'Last_day_retrieved_log.csv'
def write_to_logfile(dt, status):
    dt_log_string = dt.strftime('%Y-%m-%d')
    this_df = pd.DataFrame([{'date':dt_log_string, 'status':status}])

    try:
        log = pd.read_csv(log_path)
        log = pd.concat([log, this_df], ignore_index=True)
        log.to_csv(log_path, index=False)
    except:
        this_df.to_csv(log_path, index=False)        
            
    


critical_columns = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'PRODUCTION CODE', 'QTY','SHAPE', 'LABOR CODE', 'MAIN MEMBER', 'TOTAL MANHOURS', 'WEIGHT']


if os.path.exists(log_path):
    # open up the log 
    log = pd.read_csv(log_path)
    try:
        # get the last time a file was the status
        last_successful_date = log[log['status'] == 'Successfully completed all files'].iloc[-1]['date']
    except Exception:
        last_successful_date = '2021-01-13' # this is the day before the EVA hours started showing up in X -drive
 
else:    
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
 

print(f"Found the last successfullly completed date: {last_successful_date}")


#%%



 
possible_dir = ['Y:/','X:/','\\\\192.168.50.9\\Dropbox_(CSF)']
for ii in possible_dir:
    if os.path.exists(Path(ii)):
        base_dir = Path(ii) / 'production control' / 'EVA REPORTS FOR THE DAY'
        print(f'Using the drive: "{base_dir}"')
        break
    else:
        continue



for i in range(0,delta):
    day_dt = last_successful_dt + datetime.timedelta(days=1+i)
    
    day = day_dt.strftime('%m%d%Y')
    print(day)


    write_to_logfile(day_dt, 'Started')
   
    try:
        years = os.listdir(base_dir)
    except:
        write_to_logfile(day_dt, 'Could not connect to the dropbox')
        exit()           
    
    
    year = str(day_dt.year)
    
    year_dir = base_dir / year

    months = os.listdir(year_dir)
    
    #get it as a zero padded number -> 01 or 12
    month_num = str(day_dt.month).zfill(2)
    # '%B' gets the month name
    month_name = day_dt.strftime('%B').upper()
    
    month_str = month_num + ' - ' + month_name
    
    month_dir = year_dir / month_str
    
    if not os.path.exists(month_dir):
        print(f"The directory {month_dir} does not exist")
        continue
    
    days = os.listdir(month_dir)
    
    # reread the log in each loop
    log = pd.read_csv(log_path)
    
    already_retrieved_files = pd.unique(log['status'])
    already_retrieved_files = already_retrieved_files[already_retrieved_files != 'Started']
    already_retrieved_files = already_retrieved_files[already_retrieved_files != 'Successfully completed all files']
    
    if day in days:
    
        day_dir = month_dir / day    
    
        xls_files = glob.glob(str(day_dir / '*.xls'))
        xlsx_files = glob.glob(str(day_dir / '*.xlsx'))
        excel_files = xls_files + xlsx_files
        ph = [i for i in excel_files if 'PH' in i]
        if len(ph):
            print(ph)
        excel_files = [i for i in excel_files if not 'PH' in i]
        
        for excel_file in excel_files:
            # xls_file = [i for i in xls_files if 'T084' in i][0]
            excel_file_exists_with_previous_date = False
            # check to see if we have already recorded the xls file before
            if excel_file in already_retrieved_files:
                dates_for_file = log[log['status'] == excel_file]
                # set the indicator true 
                excel_file_exists_with_previous_date = True
                # check to see if we have the same date for that file
                if dates_for_file[dates_for_file['date'] == day_dt.strftime('%Y-%m-%d')].shape[0]:
                    # skip this file if it already exists with that date
                    continue
            
            basename = os.path.basename(excel_file)
            
            basename_components = basename.split('-')
            
            if len(basename_components) != 3:
                print('ERROR THE FILENAME IS INVALID: {}'.format(basename))
                ''' SEND ERROR THAT THE FILENAME IS INVALID
                1) Emad
                2) mildred tong
                3) me
                '''
                try:
                    basename_components = re.split(r'-|\ ', basename)
                    job_str = basename_components[0]
                    lot = basename_components[1]
                    shop = basename_components[2][:-4]
                    print(f'Able to resolve the invalid filename to its parts:\n\t{job_str}-{lot}-{shop}')
                except:
                    continue
                
            else:
                # job # is the start of the filename
                job_str = basename_components[0]
                # lot is the middle portion of the filename
                lot = basename_components[1]
                # chop off the '.xls' and only maintain the shop 
                shop = basename_components[2].split('.')[0]
                
                
            # print(basename, day, month_str, year)
            
            ''' Try to massage the lot name '''
            if lot[0].lower() == 't' and len(lot) == 4:
                print('Rule 0: {}'.format(lot))
                
            elif lot[0].lower() == 't' and len(lot) == 3:
                old_lot = lot
                lot = 'T' + lot[1:].zfill(3)
                print('Rule 1: {} to {}'.format(old_lot, lot))
                
            elif lot.isnumeric():
                old_lot = lot
                lot = 'T' + lot.zfill(3)
                print('Rule 2: {} to {}'.format(old_lot, lot))
                
            elif lot[0].lower() == 't' and lot[1:4].isnumeric():
                old_lot = lot
                lot = lot[:4]
                print('Rule 3: {} to {}'.format(old_lot, lot))
                
            # this could get TKT### or TK###
            elif lot[0:2].lower() == 'tk':
                old_lot = lot
                
                # convert TK into TKT
                if not lot.lower().startswith('tkt'):
                    split_on_number = re.split(r'(\d+)', lot)
                    split_on_number[0] = 'TKT'
                    lot = ''.join(split_on_number)
                    print(f'Rule TKT 1: {old_lot} to {lot}')
                    
                
                # make sure the number part is 3 digits 
                split_on_number = re.split(r'(\d+)', lot)
                if len(split_on_number[1]) < 3:
                    split_on_number[1] = split_on_number[1].zfill(3)
                    old_lot = lot
                    lot = ''.join(split_on_number)
                    print(f'Rule TKT 2: {old_lot} to {lot}')
                    
                # no changes were made
                if lot == old_lot:
                    print(f'Rule TKT 0: {lot}')

                
            else:
                print(f'This lot does not fill other criteria\n{lot}\t{excel_file}')
                write_to_logfile(day_dt, f'{lot} is does not match expected criterias: {excel_file}')
                # 5 + '5'
                # raise Exception('This lot does not fill other criteria')
                
            
            # we are going to iterate thru the sheet_names as numbers to find the right one b/c calling by name results in a lot of failures
            sheet_num = 0
            max_sheet_num = 6
            while sheet_num <= max_sheet_num:
                try:
                    # open the current lots file -> fingers crossed the header is alwasy row 2
                    # only maintain those 9 columns that we actually need to pass along -> smaller file sizes
                    excel_lot = pd.read_excel(excel_file, header=2, engine='xlrd', sheet_name=sheet_num, usecols=critical_columns)
                    # if for some reason it pulls it but the dataframe is shape (0,0)
                    if not excel_lot.size:
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
            excel_main_name = Path('c://downloads//') / (job_str + '.xlsx'  )
            # this is if the storage file in c:/downloads exists for that job
            if os.path.exists(excel_main_name):  
                # start by assuming headers are on row 0
                header_num = 0
                # open just the first col with 5 rows
                try:
                    excel_main_test = pd.read_excel(excel_main_name, engine='xlrd', header=header_num, nrows=5, usecols=[0])
                except:
                    excel_main_test = pd.read_excel(excel_main_name, engine='openpyxl', header=header_num, nrows=5, usecols=[0])
             
                # if the header is not 'JOB NUMBER' iterate header_num and try again
                while excel_main_test.columns[0] != 'JOB NUMBER':
                    header_num += 1
                    try:
                        excel_main_test = pd.read_excel(excel_main_name, engine='xlrd', header=header_num, nrows=5, usecols=[0])
                    except:
                        excel_main_test = pd.read_excel(excel_main_name, engine='openpyxl', header=header_num, nrows=5, usecols=[0])
    
                #open the main file with the correct header number
                try:
                    excel_main = pd.read_excel(excel_main_name, engine='xlrd', header=header_num)
                except:
                    excel_main = pd.read_excel(excel_main_name, engine='openpyxl', header=header_num)
                
                # add the lot to the df
                excel_lot['LOT'] = lot
                
                # if that LOT already exists in excel_main
                if excel_main[excel_main['LOT'] == lot].shape[0]:
                    # get rid of that LOT's records from excel_main
                    excel_main = excel_main[excel_main['LOT'] != lot]
                
                excel_lot_grouped = excel_lot.groupby(['JOB NUMBER','SEQUENCE','PAGE','MAIN MEMBER','PRODUCTION CODE','SHAPE','LABOR CODE','LOT']).sum()
                excel_lot_grouped = excel_lot_grouped.reset_index(drop=False)
                #
                    
                # append the lot's pieces to that main file 
                # excel_main = excel_main.append(excel_lot)
                excel_main = pd.concat([excel_main, excel_lot])
                try:
                    # send back to excel
                    excel_main.to_excel(excel_main_name , index=False)
                    print('Successfully added {0} to {1}'.format(basename, excel_main_name))
                except:
                    print(f"could not write {excel_main_name}")
                
                
            
            #if the file for that job DOES NOT EXIST, create it & create the excel_main variable as that LOTS data
            else:
                # add the lot to the df
                excel_lot['LOT'] = lot
                
                excel_lot.to_excel(excel_main_name, index=False)
                
                excel_main = excel_lot.copy()
            
            
            # all_LOTs_file = open(all_LOTs_imported_path,"a")
            # all_LOTs_file.write(excel_file + ", " + datetime.datetime.now().strftime('%m/%d/%Y %H:%M') + " \n")
            # all_LOTs_file.close()
            
            write_to_logfile(day_dt, excel_file)
    write_to_logfile(day_dt, 'Successfully completed all files')

