# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:37:10 2023

@author: CWilson
"""

from Pull_Fablisting_data import get_fablisting_plus_model_summary
from Pull_TimeClock_data import get_timeclock_summary
from Predictor import get_prediction_formula
from Post_to_GoogleSheet import post_observation, post_predictor
# this will be the controller for automation?


fablisting_summary = get_fablisting_plus_model_summary()
timeclock_summary = get_timeclock_summary()
predictor = None

post_observation(fablisting_summary, timeclock_summary)

# if predictor != None:
#     predictor = get_prediction_formula()
#     post_predictor()