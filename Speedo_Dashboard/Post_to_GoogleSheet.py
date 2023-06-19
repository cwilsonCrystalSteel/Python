# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:36:21 2023

@author: CWilson
"""
import sys
sys.path.append('C:\\Users\\cwilson\\documents\\python\\Speedo_Dashboard')
from production_dashboards_google_credentials import init_google_sheet
import pandas as pd
import datetime
import time
from Predictor import get_prediction_dict
import numpy as np
import os
# this one will send TimeClock & Fablisting data to google sheet





# sheet_name = 'CSM'
google_sheet_info = {'sheet_key':'1RZKV2-jt5YOFJNKM8EJMnmAmgRM1LnA9R2-Yws2XQEs',
                     'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-e051ae12d1ef.json'
                     }

def remove_empty_rows_between_data(worksheet):
        
    # get all values in worksheet
    values = worksheet.get_all_values()    
    
    empty_rows = []
    
    # iterate through rows and find empty rows
    for i, row in enumerate(values):
        if all(cell == '' for cell in row):
            empty_rows.append(i+1) # add 1 to index to match Google Sheet indices
    
    # iterate through empty rows in reverse order and delete them
    for row_index in reversed(empty_rows):
        worksheet.delete_rows(row_index)

def get_gspread_worksheet(sheet_name, sheet_key = google_sheet_info['sheet_key'], json_file = google_sheet_info['json_file']):
    sh = init_google_sheet(sheet_key, json_file)
    worksheet = sh.worksheet(sheet_name)
    remove_empty_rows_between_data(worksheet)
    worksheet = sh.worksheet(sheet_name)
    return worksheet


def get_google_sheet_as_df(shop=None, worksheet=None):
    if worksheet == None:
        worksheet = get_gspread_worksheet(shop)
        
    worksheet_list_of_lists = worksheet.get_all_values()
    df = pd.DataFrame(worksheet_list_of_lists[1:], columns=worksheet_list_of_lists[0])

    try:
        df['Timestamp'] = pd.to_datetime(df['Timestamp'])
    except Exception:
        print('could not convert timestamp to dataframe')
        
    df['IsReal'] = df['IsReal'].astype(int).astype(bool)   
    
    df.iloc[:,3:] = df[df.columns[3:]].astype(float)
    
    return df


def post_observation(gsheet_dict, isReal=True, sheet_name='CSM'):
    now = datetime.datetime.now()
    now_str = now.strftime('%m/%d/%Y %H:%M')
    
    worksheet = get_gspread_worksheet(sheet_name)
    df = get_google_sheet_as_df(worksheet=worksheet)
    
    
    # check to see if a formula exists
    df_pred_row = df[~df['IsReal']]
    # overwrite the predictor row
    if df_pred_row.shape[0]:
        idx_num = df_pred_row.index[-1]
        row_num = idx_num + 2
    else:
        row_num = df.index[-1] + 3  
        
    for i,col in enumerate(df.columns):
        colletter = chr(ord('A') + i)
        if col == 'Timestamp':
            value = now_str
        elif col == 'IsReal':
            value = int(isReal)
        elif col == 'Total Hours':
            value = (gsheet_dict['Direct Hours'] + gsheet_dict['Indirect Hours']).round(2)
        elif col == 'Efficiency':
            if gsheet_dict['Direct Hours']:
                value = (gsheet_dict['Earned Hours'] / gsheet_dict['Direct Hours']).round(2)
            else:
                value = 0
        else:
            value = gsheet_dict[col]
        
        cell = colletter + str(row_num)
        print(cell, col, value)
        time.sleep(1)
        if isinstance(value, np.int32):
            value = int(value)
        try:
            worksheet.update(cell, value, value_input_option='USER_ENTERED')
        except:
            num_columns = len(df.columns)-1
         
            worksheet.append_row([''] * num_columns, value_input_option='USER_ENTERED')
            worksheet.update(cell, value, value_input_option='USER_ENTERED')
    
    # then paste the observed data
    return None



def post_predictor():
   
    df = get_google_sheet_as_df()
    
    gsheet_pred_dict = get_prediction_dict(df)
    
    post_observation(gsheet_pred_dict, isReal=False)
    
    
    
    
    # simply append a isReal=0 row
    return None







def move_to_archive(shop=None):
    
    archive_file = 'c:\\users\\cwilson\\documents\\python\\speedo_dashboard\\archive_' + shop + '.csv'
    
    if os.path.exists(archive_file):
        # archive = pd.read_csv(archive_file, index_col=0)
        
        worksheet = get_google_sheet_as_df(shop, worksheet=None)
        # only archive if we have data & we have over 100 rows
        if worksheet.shape[0] and worksheet.shape[0] > 90:
            # get the last row to archive
            to_archive = worksheet.iloc[:1,:]
            
            with open(archive_file, 'a', newline='') as f:
                to_archive.to_csv(f, header=False, index=False, line_terminator='\n')
            print('row appened to archive csv file')
                
            sh = init_google_sheet(google_sheet_info['sheet_key'], google_sheet_info['json_file'])
            # Get the first worksheet
            worksheet = sh.worksheet(shop)
            
            # Delete the second row
            worksheet.delete_rows(2)
            print('row 2 deleted from {} google sheet'.format(shop))
        
        
    else:
        print('create the archive')
        worksheet = get_google_sheet_as_df()
        worksheet = worksheet.iloc[:worksheet.shape[0]-100]
        worksheet.to_csv(archive_file, index=False)
        
    
    
    return None


def get_shop_b_jobs():
    # URL: https://docs.google.com/spreadsheets/d/1SyV3tqt7a0-tFRNWYNtmNajzjJ4KY9JC4isgHUkj7WU/edit#gid=616322759
    # key = from d/ xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx /edit
    weeklyshophourssheetkey = '1SyV3tqt7a0-tFRNWYNtmNajzjJ4KY9JC4isgHUkj7WU'

    # get the sheet as a model.worksheet
    jobsummary_worksheet = get_gspread_worksheet(sheet_name = 'Job Summary', sheet_key = weeklyshophourssheetkey)
    # convert to list of lists
    jobsummary_list = jobsummary_worksheet.get_all_values()
    # convert to data frame
    jobsummary_df = pd.DataFrame(jobsummary_list[1:], columns=jobsummary_list[0])
    # keep only things from the Shop B title 'TN Misc Metals Productive Hours' onwards
    shopB_start = list(jobsummary_df.columns).index('TN Misc Metals Productive Hours')
    # get only from where the TN misc metals is and to the right. also only get from row 9 down
    shopB = jobsummary_df.iloc[8:, shopB_start:]
    # reset index
    shopB = shopB.reset_index(drop=True)
    # set columns to be the first row
    shopB.columns = shopB.iloc[0]
    # get rid of the columns row now
    shopB = shopB.iloc[1:]
    # get a list of the job codes
    jobs_to_exclude = pd.unique(shopB['Job Code'])
    # get rid of any empty values
    jobs_to_exclude = [i for i in jobs_to_exclude if i != '']
    # get the numbers only in case there are jobs with names
    jobs_numbers_only = [i.split(' ')[0] for i in jobs_to_exclude]
    # append the numbers to the original list
    jobs_to_exclude += jobs_numbers_only
    # get only unique occurences & back to a list again
    jobs_to_exclude = list(set(jobs_to_exclude))
    # return
    return jobs_to_exclude