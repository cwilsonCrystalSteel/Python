# -*- coding: utf-8 -*-
"""
Created on Sat Mar 11 08:53:58 2023

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')

from sklearn.linear_model import LinearRegression
from Post_to_GoogleSheet import get_google_sheet_as_df
import datetime
import numpy as np
import pandas as pd
# this will be the function that does prediction 
now = datetime.datetime.now()


start_of_workday = now.date()
if now.hour < 6:
    start_of_workday += datetime.timedelta(day=-1)

start_of_workday = datetime.datetime.combine(start_of_workday, datetime.time.min)
start_of_workday = start_of_workday.replace(hour=6)


def get_seconds_until_end_of_workday(start):
    end_of_workday = start_of_workday + datetime.timedelta(days=1)
    seconds = (end_of_workday - start).total_seconds()    
    return seconds


def get_prediction_dict():
    
    df = get_google_sheet_as_df()
    today_df = df[df['Timestamp'] >= start_of_workday]
    
    start = min(today_df['Timestamp'])
    prediction_value = get_seconds_until_end_of_workday(start)
    
    x = np.array((today_df['Timestamp'] - start)) / np.timedelta64(1, 's')
    X = np.array([np.ones(x.size),x]).transpose()
    y = np.array(today_df['Direct Hours'])
    
    regression = LinearRegression()
    regression.fit(X,y)
    regression.predict(X)    
    return None