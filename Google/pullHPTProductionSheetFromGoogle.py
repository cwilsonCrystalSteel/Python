# -*- coding: utf-8 -*-
"""
Created on Sat Apr  5 12:29:34 2025

@author: Netadmin
"""

from utils.attendance_google_sheets_credentials_startup import init_google_sheet as init_google_sheet_production_worksheet
import pandas as pd
import re

_ProductionWorksheetGooglekey = "1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4"
sh = init_google_sheet_production_worksheet(_ProductionWorksheetGooglekey)
col_row = 3

def pullProductionSheetFromGoogle(drop_nonnumeric=True):

    
    # get the values from the shipping schedule as a list of lists
    worksheet = sh.worksheet('Production Sheet').get_all_values()
    
    df = pd.DataFrame(worksheet[col_row + 1:], columns=worksheet[col_row])    

    df = df.iloc[:,:15]
    
    hardcoded_columns = ['Job', 'Hrs/ ton', 'Seq. #', 'Tonnage', 'Req. Hrs.', 'Tons for Fab',
           'Hours Worked', 'Tons Completed', 'Remaining Hrs.', 'Booked hours',
           'Variance should be 0', 'Work on hand', 'Complete', 'Productivity ',
           'Shop']
    
    df.columns = hardcoded_columns
    
    add_new_line_text = 'Add new lines above here'
    try:
        add_new_line_index = list(df['Seq. #']).index(add_new_line_text)
    except ValueError:
        for i in df.columns:
            try:
                add_new_line_index = list(df[i]).index(add_new_line_text)
                break
            except:
                continue
    
    # shorten the dataframe by the value of the new line text
    df = df.iloc[:add_new_line_index, :]
    
    # now we can get rid of rows that are missing vlaues in the job #
    df = df[df['Job'] != '']
    
    df = df[['Job','Hrs/ ton','Shop']]
    
    df['Job'] = df['Job'].apply(lambda x: re.sub(r'[^0-9]', '', x.upper()))
    
    df['Shop'] = df['Shop'].apply(lambda x: x.strip())
    
    df['Hrs/ ton'] = df['Hrs/ ton'].apply(lambda x: re.sub(r'[^0-9.]','', x))

    df['Job'] = pd.to_numeric(df['Job'], errors='coerce')
    df['Hrs/ ton'] = pd.to_numeric(df['Hrs/ ton'], errors='coerce')
    
    df = df.rename(columns={'Hrs/ ton':'hpt'})
    
    return df
    
