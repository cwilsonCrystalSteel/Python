# -*- coding: utf-8 -*-
"""
Created on Mon Apr  7 09:46:36 2025

@author: Netadmin
"""


import datetime 
import pandas as pd
import numpy as np
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2
from get_model_estimate_hours_attached_to_fablisting_SQL import apply_model_hours_SQL, call_to_insert
from utils.google_sheets_credentials_startup import init_google_sheet
from insertFablistingToSQL import lotnumber_cleaner


#%%

start_date = '04/06/2025'
start_dt = datetime.datetime.strptime(start_date, '%m/%d/%Y')

end_date = '04/07/2025'
end_dt = datetime.datetime.strptime(end_date, '%m/%d/%Y')

sheet = 'FED QC Form'
shop = sheet[:3]

fablisting = grab_google_sheet(sheet, start_date, end_date, start_hour='use_function', include_sheet_name=True)
call_to_insert(fablisting)
# eva = apply_model_hours_SQL(how='best', keep_diagnostic_cols=True)
eva = apply_model_hours_SQL(how=['eva_hours_lotslog','eva_hours','hpt_hours'], keep_diagnostic_cols=True)

#%%
compared_to_sheet_key = '1gTBo9c0CKFveF892IgWEcP2ctAtBXoI3iqjEvZVtl5k'
compared_to_sheet_name = 'Form Responses'
sh = init_google_sheet(compared_to_sheet_key)

# open the QC input sheet
worksheet = sh.worksheet(compared_to_sheet_name)
# convert all values to a list of lists
all_values = worksheet.get_all_values()
#%%
# only takes the last 100 rows as data for the dataframe
comp = pd.DataFrame(columns=all_values[1], data=all_values[2:])

comp['lotcleaned'] = comp['Lot #'].apply(lotnumber_cleaner)
comp['Date'] = pd.to_datetime(comp['Date'])

comp = comp.rename(columns={'Piecemark':'Piece Mark - REV'})

comp['Job #'] = pd.to_numeric(comp['Job #'], errors='coerce')

comp['EVA'] = pd.to_numeric(comp['EVA'])

justin = comp[(comp['Date'] >= start_dt) & (comp['Date'] <= end_dt) & (comp['Site'] == shop)]
#%%



columns_for_pcmark_comp = ["Job #",'lotcleaned','Piece Mark - REV']
my_eva_cols = ['eva_hours_dropbox', 'eva_hours_lotslog',
               'eva_hours_lotslogjobaverage', 'evaperpound_lotslog']


by_pcmark = pd.merge(left = justin[columns_for_pcmark_comp + ['EVA']],
                     right = eva[columns_for_pcmark_comp + ['Earned Hours'] + my_eva_cols],
                     left_on = columns_for_pcmark_comp,
                     right_on = columns_for_pcmark_comp,
                     how = 'outer',
                     indicator=True)

by_pcmark = by_pcmark.sort_values(by='_merge')

by_pcmark['diff'] = by_pcmark['EVA'] - by_pcmark['eva_hours_lotslog']
by_pcmark['evaperton_lotslog'] = by_pcmark['evaperpound_lotslog'] * 2000
