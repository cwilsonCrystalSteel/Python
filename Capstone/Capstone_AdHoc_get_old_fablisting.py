# -*- coding: utf-8 -*-
"""
Created on Mon Sep 18 09:22:37 2023

@author: CWilson
"""
import datetime 
import pandas as pd
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2



for year in [2022,2021,2020]:
    
    
    start_dt = datetime.datetime(year=year, month=1, day=1, hour=1, minute=0, second=0)
    start_date = start_dt.strftime('%m/%d/%Y')
    end_dt = start_dt + datetime.timedelta(days=365)
    end_date = end_dt.strftime('%m/%d/%Y')
    start_date = start_dt.strftime('%m/%d/%Y')
    # format as the date string
    end_date = end_dt.strftime('%m/%d/%Y')
    
    for shop in ['CSM','CSF','FED']:
        print('Pulling fablisting for {}: {} to {}'.format(shop, start_date, end_date))
        
        sheet_name = shop + ' ' + str(year)
        # get fablisting for all of start_date and all of end_date
        fablisting = grab_google_sheet(sheet_name, start_date, end_date, sheet_key="1nyM3ocSMM4EeMrOmmh0UBEDGUM6L8lte5TPYXcLPa2s")
        # get dates between yesterday at 6 am and today at 6 am
        fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
        # get the model hours attached to fablisting
        with_model = apply_model_hours2(fablisting, how = 'model but Justins dumb way of getting average hours', fill_missing_values=True, shop=shop)
        # with_model = apply_model_hours2(fablisting, how = 'model', fill_missing_values=False, shop=sheet[:3])

        with_model.to_csv(".\\" + sheet_name + '.csv')







