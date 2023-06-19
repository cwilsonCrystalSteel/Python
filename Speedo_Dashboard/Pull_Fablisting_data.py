# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:36:02 2023

@author: CWilson
"""

import datetime 
import pandas as pd
import numpy as np
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2

# this one will pull the fablisting data
state = 'TN'
sheet = 'CSM QC Form' 

# start_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
# start_date = start_dt.strftime('%m/%d/%Y')
# end_dt = start_dt + datetime.timedelta(days=1)
# end_date = end_dt.strftime('%m/%d/%Y')

def get_fablisting_plus_model_summary(start_dt, end_dt, sheet, exclude_jobs_list=None):
    
    start_date = start_dt.strftime('%m/%d/%Y')
    # format as the date string
    end_date = end_dt.strftime('%m/%d/%Y')
    print('Pulling fablisting for: {} to {}'.format(start_date, end_date))
    # get fablisting for all of start_date and all of end_date
    fablisting = grab_google_sheet(sheet, start_date, end_date)
    # get dates between yesterday at 6 am and today at 6 am
    fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
    # get the model hours attached to fablisting
    with_model = apply_model_hours2(fablisting, how = 'model but Justins dumb way of getting average hours', fill_missing_values=True, shop=sheet[:3])
    # with_model = apply_model_hours2(fablisting, how = 'model', fill_missing_values=False, shop=sheet[:3])

    '''
    with_model['date'] = with_model['Timestamp'].dt.date
    with_model.to_excel('c:\\users\\cwilson\\downloads\\fablisting_with_model.xlsx')
    with_model.groupby('date').sum().to_excel('c:\\users\\cwilson\\downloads\\fablisting_with_model_date_grouped.xlsx')
    
    
    pieces = with_model.groupby(['Job #','Lot #', 'Piece Mark - REV','Has Model']).agg({'Weight':sum, 'Quantity':sum, 'Earned Hours':sum, 'Timestamp':lambda x: ', '.join(x.dt.strftime('%m/%d/%Y %H:%M'))})
    # go to the form responses tab of fablisting & just copy paste the rows you want (with header0 into excel file)
    formresponses = pd.read_excel('c:\\users\\cwilson\\downloads\\formResponses.xlsx')
    fr_pieces = formresponses[formresponses['Site'] == 'CSM'].groupby(['Job #','Lot #','Piecemark']).agg({'Weight':sum, 'Qty':sum, 'EVA':sum, 'Date':lambda x: ', '.join(x.dt.strftime('%m/%d/%Y'))})
    pieces = pieces.reset_index()
    fr_pieces = fr_pieces.reset_index()
    pieces['Lot #'] = pieces['Lot #'].astype(str)
    fr_pieces['Lot #'] = fr_pieces['Lot #'].astype(str)
    joined = pd.merge(pieces.reset_index(), fr_pieces.reset_index(), left_on=['Job #','Lot #','Piece Mark - REV'], right_on=['Job #','Lot #','Piecemark'])
    joined = joined[['Job #','Lot #','Piecemark','Has Model','Weight_x','Weight_y','Quantity','Qty','Earned Hours','EVA','Timestamp','Date']]
    joined = joined.rename(columns={'Quantity':'Qty_x','Qty':'Qty_y','Earned Hours':'EVA_x','EVA':'EVA_y','Date':'FR Date'})
    joined['Diff'] = joined['EVA_x'] - joined['EVA_y']
    joined.to_excel('c:\\users\\cwilson\\downloads\\withModel_vs_formResponses.xlsx')
    '''
    if exclude_jobs_list != None:
        with_model = with_model[~with_model['job #'].isin(exclude_jobs_list)]
    
    num_with_model = with_model['Has Model'].sum()
    num_without_model = with_model.shape[0] - num_with_model
    
    earned_hours = np.round(with_model['Earned Hours'].sum(), 2)
    tonnage = np.round((with_model['Weight'].sum() / 2000), 2)
    quantity = int(with_model['Quantity'].sum())
    
    
    return {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity}