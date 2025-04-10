# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:36:02 2023

@author: CWilson
"""

import datetime 
import pandas as pd
import numpy as np
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2
from get_model_estimate_hours_attached_to_fablisting_SQL import apply_model_hours_SQL, call_to_insert

# this one will pull the fablisting data
state = 'TN'
sheet = 'CSM QC Form' 

# start_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
# start_dt = datetime.datetime(year=2023, month=1, day=1, hour=1, minute=0, second=0)
# start_date = start_dt.strftime('%m/%d/%Y')
# end_dt = start_dt + datetime.timedelta(days=365)
# end_date = end_dt.strftime('%m/%d/%Y')

def get_fablisting_plus_model_summary(start_dt, end_dt, sheet, output_fablisting_copy=False, exclude_jobs_dict=None):
    
    start_date = start_dt.strftime('%m/%d/%Y')
    # format as the date string
    end_date = end_dt.strftime('%m/%d/%Y')
    print('Pulling fablisting for: {} to {}'.format(start_date, end_date))
    # get fablisting for all of start_date and all of end_date
    fablisting = grab_google_sheet(sheet, start_date, end_date)
    # get dates between yesterday at 6 am and today at 6 am
    fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
    # get the model hours attached to fablisting
    call_to_insert(fablisting, sheet, source='Pull_Fablisting_data')
    with_model = apply_model_hours_SQL(how='best', keep_diagnostic_cols=False)
    # with_model = apply_model_hours2(fablisting, how = 'model but Justins dumb way of getting average hours', fill_missing_values=True, shop=sheet[:3])
    # with_model = apply_model_hours2(fablisting, how = 'model', fill_missing_values=False, shop=sheet[:3])


    if exclude_jobs_dict != None:
        # get the jobs to exclude from the dict and shop name
        excluded_jobs = exclude_jobs_dict[sheet[:3]]
        # remove those pieces
        with_model = with_model[~with_model['Job #'].isin(excluded_jobs)]
    
    # count number of peices with model horus
    # num_with_model = with_model['Has Model'].sum()
    # count pieces without model  hours
    # num_without_model = with_model.shape[0] - num_with_model
    # sum the number of earned hours
    earned_hours = np.round(with_model['Earned Hours'].sum(), 2)
    # sum the weight into tons
    tonnage = np.round((with_model['Weight'].sum() / 2000), 2)
    # get the quantity of pieces
    quantity = int(with_model['Quantity'].sum())
    
    if output_fablisting_copy:
        return {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity, 'Fablisting':with_model}
    else:
        return {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity}