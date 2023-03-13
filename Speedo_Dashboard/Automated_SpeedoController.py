# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:37:10 2023

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
from Pull_Fablisting_data import get_fablisting_plus_model_summary
from Pull_TimeClock_data import get_timeclock_summary
from Predictor import get_prediction_formula
from Post_to_GoogleSheet import post_observation, post_predictor

# this will be the controller for automation?


fablisting_summary = get_fablisting_plus_model_summary()
timeclock_summary = get_timeclock_summary()

gsheet_dict = {}
gsheet_dict.update(fablisting_summary)
gsheet_dict.update(timeclock_summary)

post_observation(gsheet_dict)

predictor = None


# if predictor != None:
#     predictor = get_prediction_formula()
#     post_predictor()