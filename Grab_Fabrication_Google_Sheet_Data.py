# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 11:30:32 2021

@author: CWilson
"""

import pandas as pd
import gspread
from utils.google_sheets_credentials_startup import init_google_sheet
import datetime 

sheet_key="1gTBo9c0CKFveF892IgWEcP2ctAtBXoI3iqjEvZVtl5k"

# daily_fab_listing_google_sheet_key = "1gTBo9c0CKFveF892IgWEcP2ctAtBXoI3iqjEvZVtl5k"
def convert_hours_to_dayshift(start_date, end_date, start_hour=6):
        start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
        end_dt = datetime.datetime.strptime(end_date, "%m/%d/%Y")
        arg_start_dt = start_dt.replace(hour=start_hour, minute=0, second=0)
        
        arg_end_dt = end_dt + datetime.timedelta(days=1)
        arg_end_dt = arg_end_dt.replace(hour=start_hour, minute=0, second=0)
        
        print(f"{start_date} : {arg_start_dt.strftime('%m/%d/%Y %H:%M')}")
        print(f"{end_date} : {arg_end_dt.strftime('%m/%d/%Y %H:%M')}")
        
        return arg_start_dt, arg_end_dt
    

def grab_google_sheet(sheet_name, 
                      start_date="04/01/2025", end_date="04/01/2025", 
                      start_hour=0,
                      include_sheet_name=False,
                      include_hyperlink=False):
    
    if isinstance(start_hour, int):
        # set start date to datetime
        start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
        # make it so the start time on the start date is 4:00 am.... prevent any potential overlap from 2nd shift
        start_dt = start_dt.replace(hour=start_hour, minute=0, second=0)
        # set end date to datetime
        end_dt = datetime.datetime.strptime(end_date, "%m/%d/%Y")
        # set end date to be at 23:59
        end_dt = end_dt.replace(hour=23, minute=59, second=59)
    
    elif start_hour == 'use_function':
        # justin says that the start of dayshift is 6 am
        # a day runs from 6 am to 6 am 
        start_dt, end_dt = convert_hours_to_dayshift(start_date, end_date, 6)
        
        
    
    # key for the daily fab listing google sheet
    
    # open the daily fab listing google sheet
    sh = init_google_sheet(sheet_key)


    
    # open the QC input sheet
    worksheet = sh.worksheet(sheet_name)
    # convert all values to a list of lists
    all_values = worksheet.get_all_values()
    
    # adding hyperlinks to the 'last column' of the sheet
    if include_hyperlink:
        # chatGPT
        sheet_metadata = sh.fetch_sheet_metadata()
        worksheet_gid = None
        for sheet in sheet_metadata['sheets']:
            if sheet['properties']['title'] == sheet_name:
                worksheet_gid = sheet['properties']['sheetId']
                break
        
        # make a lambda function to take in rownumber and transform to URL
        make_col_a_url = lambda rownum: (
            'hyperlink' if rownum == 0
            else f"https://docs.google.com/spreadsheets/d/{sheet_key}/edit#gid={worksheet_gid}&range=A{rownum+1}"
            )
        # addd the url as the last item in each list
        all_values = [rowlist + [make_col_a_url(rownum)] for rownum, rowlist in enumerate(all_values)]

    
    # columns are first row
    # data starts on Google Sheets Row Number = 3 (python 2)
    df = pd.DataFrame(columns=all_values[0], data=all_values[2:])
    # this is to force columns incase any accidentaly renaming happens on the google sheet
    if 'Timestamp' not in df.columns:
        try:
            columns = ['', 'Timestamp', 'Job #', 'Lot #', 'Quantity', 
                       'Piece Mark - REV', 'Weight', 'Fitter', 'Fit QC', 
                       'Welder', 'Weld QC', 'Does This Piece Have a Defect?', 
                       'Date Found', 'Sequence Number', 'Quantity', 'Member Type'] 
                       # 'Worked By', 'Inspected By', 'Description of Non-Conformance', 
                       # 'Defect Category', 'Defect Sub Category', 'Defect Type', 
                       # 'Disposition', 'Repaired By', 'Repair Inspected By', 
                       # 'Rework Time', '', '']
            # we want to fill in the blanks with empty colnames
            if len(columns) < df.shape[1]:
                columns = columns + [''] * (df.shape[1] - len(columns))
                # change the last name to hyperlink if using it
                if include_hyperlink:
                    columns[-1] = 'hyperlink'
            # now fix up the dataframe with consistent and reliable column names
            df = pd.DataFrame(columns = columns, data=all_values)
        except:
            return pd.DataFrame(data=all_values)
    # Convert entry time from string to datetime
    # the errors='coerce' makes it so times that don't conform to a datetime will be changed to NaT value
    df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce')
    # remove any dates that became NaT values from above step
    df = df[df['Timestamp'].notna()]
    
    # get only pieces after start date
    df = df[df['Timestamp'] >= start_dt]
    # get only pieces before end date. this is a catch just in case
    df = df[df['Timestamp'] <= end_dt]
    # Removes duplicated column
    df = df.loc[:,~df.columns.duplicated()]
    if include_hyperlink:
        df = df[list(df.columns[:12]) + ['hyperlink']]
    else:
        # Drop the columns regarding defects, keeps the isDefect yes/no
        df = df[df.columns[:12]]
    # Resets the index, don't need old index
    df = df.reset_index(drop=True)
    # only keep the first 4 digits in the job #
    df['Job #'] = df['Job #'].astype(str).str[:4]
    
    
    
    # This trys to convert the 'Job #', 'Quantity', "Weight' columns to a number
    # if it cannot convert it to a number, it sets it as NaN
    df['Job #'] = df['Job #'].apply(pd.to_numeric, errors='coerce', downcast='integer')
    df['Quantity'] = df['Quantity'].apply(pd.to_numeric, errors='coerce')
    df['Weight'] = df['Weight'].apply(pd.to_numeric, errors='coerce')
    
    # df[df.columns[[2,3,5]]] = df[df.columns[[2,3,5]]].apply(pd.to_numeric, errors='coerce')
   
    # removes the row if there is a NaN in the 'Job #' column
    df = df[~df["Job #"].isna()]
    # removes the row if there is a NaN in the 'Quantity' column
    df = df[~df["Quantity"].isna()]
    # removes the row if there is a NaN in the 'Weight' column
    df = df[~df["Weight"].isna()]
    
    if include_sheet_name:
        df['sheetname'] = sheet_name
    

    return df