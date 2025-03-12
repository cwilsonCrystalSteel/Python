# -*- coding: utf-8 -*-
"""
Created on Thu Aug 26 08:58:05 2021

@author: CWilson
"""

import sys
sys.path.append('c:\\users\\cwilson\\documents\\python\\Lots_schedule_calendar\\')
sys.path.append('c://users//cwilson//documents//python//Attendance Project//')

import datetime
import pandas as pd
from utils.attendance_utils.google_sheets_credentials_startup import init_google_sheet


sh = init_google_sheet("1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4")
# get the values from the shipping schedule as a list of lists
worksheet = sh.worksheet('Shipping Sched.').get_all_values()
# convert to a dataframe. row 2 is columns, dont care about stuff before row 10
ss = pd.DataFrame(worksheet[10:-8], columns=worksheet[2])
ss = ss[ss['Job'] != '2123']
# empty dict to get rid of line breaks in the column names
new_cols = {}
for col in ss.columns:
    new_col = col.replace('\n', ' ')
    new_cols[col] = new_col
    
# replace columns with new columns w/o line breaks
ss = ss.rename(columns=new_cols)

# get the ones without type of work
no_tow = ss[ss['Type of Work'] == '']
no_tow_bad_num = no_tow[pd.to_numeric(no_tow['Number'], errors='coerce').isna()]
no_tow_bad_date = no_tow[pd.to_datetime(no_tow['Delivery'], errors='coerce').isna()]

# get the ones with the number column not actually having a number
good_tow = ss[~(ss['Type of Work'] == '')]
bad_num = good_tow[pd.to_numeric(good_tow['Number'], errors='coerce').isna()]
bad_date = good_tow[pd.to_datetime(good_tow['Delivery'], errors='coerce').isna()]


bad_num_count = bad_num.groupby(by='PM').count()['Job']
bad_num_count['Total'] = bad_num_count.sum()
bad_date_count = bad_date.groupby(by='PM').count()['Job']
bad_date_count['Total'] = bad_date_count.sum()
