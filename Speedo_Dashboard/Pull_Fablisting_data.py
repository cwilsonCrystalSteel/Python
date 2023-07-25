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
from Post_to_GoogleSheet import get_production_worksheet_job_hours

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

    try:
        no_earned_hours_1 = with_model[with_model['Earned Hours'].isna()]
        production_worksheet_hpt = get_production_worksheet_job_hours()
    
        # get horus based on shop & job
        production_worksheet_hpt_shop = production_worksheet_hpt[production_worksheet_hpt['Shop'] == sheet[:3]]
        nada_1 = pd.merge(no_earned_hours_1, production_worksheet_hpt_shop, on='Job #', how='left')
        nada_1 = nada_1.set_index(no_earned_hours_1.index)
        nada_1['Earned Hours'] = nada_1['HPT'] * nada_1['Weight']/2000
        nada_1['Hours Per Pound'] = nada_1['HPT']
        with_model.loc[nada_1.index,:] = nada_1
        # get hours based on job if any remaining ones left 
        no_earned_hours_2 = with_model[with_model['Earned Hours'].isna()]
        nada_2 = pd.merge(no_earned_hours_2, production_worksheet_hpt, on='Job #', how='left')
        nada_2 = nada_2.set_index(no_earned_hours_2.index)
        nada_2['Earned Hours'] = nada_2['HPT'] * nada_1['Weight'] / 2000
        nada_2['Hours Per Pound'] = nada_2['HPT']
        with_model.loc[nada_2.index,:] = nada_2
    except:
        print('could not go to production worksheet for more hpt valeus')

    if exclude_jobs_list != None:
        with_model = with_model[~with_model['Job #'].isin(exclude_jobs_list)]
    
    num_with_model = with_model['Has Model'].sum()
    num_without_model = with_model.shape[0] - num_with_model
    
    earned_hours = np.round(with_model['Earned Hours'].sum(), 2)
    tonnage = np.round((with_model['Weight'].sum() / 2000), 2)
    quantity = int(with_model['Quantity'].sum())
    
    
    return {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity}