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


    if exclude_jobs_list != None:
        with_model = with_model[~with_model['Job #'].isin(exclude_jobs_list)]
    
    num_with_model = with_model['Has Model'].sum()
    num_without_model = with_model.shape[0] - num_with_model
    
    earned_hours = np.round(with_model['Earned Hours'].sum(), 2)
    tonnage = np.round((with_model['Weight'].sum() / 2000), 2)
    quantity = int(with_model['Quantity'].sum())
    
    
    return {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity}