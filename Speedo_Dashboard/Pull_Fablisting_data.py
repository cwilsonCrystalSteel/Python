# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:36:02 2023

@author: CWilson
"""

import datetime 
import pandas as pd
import numpy as np
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2, fill_missing_model_earned_hours

# this one will pull the fablisting data

def get_fablisting_plus_model(sheet, start_date, end_date):
    state = 'TN'
    sheet = 'CSM QC Form' 
    start_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    start_date = start_dt.strftime('%m/%d/%Y')
    
    start_dt = datetime.datetime.strptime(start_date, '%m/%d/%Y')
    # the production day runs from Day 0 6AM to Day + 1 5:59 AM
    start_dt = start_dt.replace(hour=6)
    # add one day to make the end time start + one day
    end_dt = start_dt + datetime.timedelta(days=1)    
    # format as the date string
    end_date = end_dt.strftime('%m/%d/%Y')
    # get fablisting for all of start_date and all of end_date
    fablisting = grab_google_sheet(sheet, start_date, end_date)
    # get dates between yesterday at 6 am and today at 6 am
    fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
    # get the model hours attached to fablisting
    with_model = apply_model_hours2(fablisting, fill_missing_values=False, shop=sheet[:3])
    
    num_with_model = with_model['Has Model'].sum()
    num_without_model = with_model.shape[0] - num_with_model
    
    earned_hours = with_model[with_model['Has Model']]['Earned Hours'].sum()
    
    return with_model