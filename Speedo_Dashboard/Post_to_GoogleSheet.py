# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:36:21 2023

@author: CWilson
"""

from production_dashboards_google_credentials import init_google_sheet
import pandas as pd
import datetime
import time
from Predictor import get_prediction_formula
# this one will send TimeClock & Fablisting data to google sheet
now = datetime.datetime.now()
now_str = now.strftime('%m/%d/%Y %H:%M')




google_sheet_info = {'sheet_key':'1RZKV2-jt5YOFJNKM8EJMnmAmgRM1LnA9R2-Yws2XQEs',
                     'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-890ed2bf828b.json'
                     }
shop = 'CSM'
sh = init_google_sheet(google_sheet_info['sheet_key'], google_sheet_info['json_file'])
worksheet = sh.worksheet(shop)
worksheet_list_of_lists = worksheet.get_all_values()
df = pd.DataFrame(worksheet_list_of_lists[1:], columns=worksheet_list_of_lists[0])


def post_observation(fablisting_summary, timeclock_summary):
    # first clear the isReal=0
    summary = {}
    summary.update(fablisting_summary)
    summary.update(timeclock_summary)
    # check to see if a formula exists
    df_pred_row = df[~df['IsReal'].astype(int).astype(bool)]
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
            value = 1
        elif col == 'Total Hours':
            value = (summary['Direct Hours'] + summary['Indirect Hours']).round(2)
        else:
            value = summary[col]
        
        cell = colletter + str(row_num)
        print(cell, col, value)
        time.sleep(1)
        worksheet.update(cell, value, value_input_option='USER_ENTERED')
    
    # then paste the observed data
    return None



def post_predictor():
    # simply append a isReal=0 row
    return None