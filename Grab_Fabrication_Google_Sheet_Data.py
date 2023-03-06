# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 11:30:32 2021

@author: CWilson
"""

import pandas as pd
import gspread
from google_sheets_credentials_startup import init_google_sheet
import datetime 

def grab_google_sheet(sheet_name, start_date="03/06/1997", end_date="03/06/1997", start_hour=0):
    
    
    # key for the daily fab listing google sheet
    daily_fab_listing_google_sheet_key = "1gTBo9c0CKFveF892IgWEcP2ctAtBXoI3iqjEvZVtl5k"
    # open the daily fab listing google sheet
    sh = init_google_sheet(daily_fab_listing_google_sheet_key)


    
    # open the QC input sheet
    worksheet = sh.worksheet(sheet_name)
    # convert all values to a list of lists
    all_values = worksheet.get_all_values()
    # only takes the last 100 rows as data for the dataframe
    # df = pd.DataFrame(columns = all_values[0], data=all_values[-200:])
    df = pd.DataFrame(columns=all_values[0], data=all_values[2:])
    # Convert entry time from string to datetime
    # the errors='coerce' makes it so times that don't conform to a datetime will be changed to NaT value
    df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce')
    # remove any dates that became NaT values from above step
    df = df[df['Timestamp'].notna()]
    # set start date to datetime
    start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
    # make it so the start time on the start date is 4:00 am.... prevent any potential overlap from 2nd shift
    start_dt = start_dt.replace(hour=start_hour, minute=0, second=0)
    # set end date to datetime
    end_dt = datetime.datetime.strptime(end_date, "%m/%d/%Y")
    # set end date to be at 23:59
    end_dt = end_dt.replace(hour=23, minute=59, second=59)
    # get only pieces after start date
    df = df[df['Timestamp'] >= start_dt]
    # get only pieces before end date. this is a catch just in case
    df = df[df['Timestamp'] <= end_dt]
    # Removes duplicated column
    df = df.loc[:,~df.columns.duplicated()]
    # Drop the columns regarding defects, keeps the isDefect yes/no
    df = df[df.columns[:12]]
    # Resets the index, don't need old index
    df = df.reset_index(drop=True)
    # only keep the first 4 digits in the job #
    df['Job #'] = df['Job #'].astype(str).str[:4]
    
    
    
    # This trys to convert the 'Job #', 'Quantity', "Weight' columns to a number
    # if it cannot convert it to a number, it sets it as NaN
    df['Job #'] = df['Job #'].apply(pd.to_numeric, errors='coerce')
    df['Quantity'] = df['Quantity'].apply(pd.to_numeric, errors='coerce')
    df['Weight'] = df['Weight'].apply(pd.to_numeric, errors='coerce')
    
    # df[df.columns[[2,3,5]]] = df[df.columns[[2,3,5]]].apply(pd.to_numeric, errors='coerce')
   
    # removes the row if there is a NaN in the 'Job #' column
    df = df[~df["Job #"].isna()]
    # removes the row if there is a NaN in the 'Quantity' column
    df = df[~df["Quantity"].isna()]
    # removes the row if there is a NaN in the 'Weight' column
    df = df[~df["Weight"].isna()]
    
    

    return df