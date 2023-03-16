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

start_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
start_date = start_dt.strftime('%m/%d/%Y')
end_dt = start_dt + datetime.timedelta(days=1)
end_date = end_dt.strftime('%m/%d/%Y')

def get_fablisting_plus_model_summary(start_dt, end_dt, sheet):
    
    start_date = start_dt.strftime('%m/%d/%Y')
    # format as the date string
    end_date = end_dt.strftime('%m/%d/%Y')
    print('Pulling fablisting for: {} to {}'.format(start_date, end_date))
    # get fablisting for all of start_date and all of end_date
    fablisting = grab_google_sheet(sheet, start_date, end_date)
    # get dates between yesterday at 6 am and today at 6 am
    fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
    # get the model hours attached to fablisting
    with_model = apply_model_hours2(fablisting, fill_missing_values=False, shop=sheet[:3])
    
    num_with_model = with_model['Has Model'].sum()
    num_without_model = with_model.shape[0] - num_with_model
    
    earned_hours = with_model[with_model['Has Model']]['Earned Hours'].sum().round(2)
    tonnage = (with_model['Weight'].sum() / 2000).round(2)
    quantity = int(with_model['Quantity'].sum())
    
    
    return {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity}