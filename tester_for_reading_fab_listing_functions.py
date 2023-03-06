# -*- coding: utf-8 -*-
"""
Created on Mon Dec 14 15:18:27 2020

@author: CWilson
"""

"used to test out the read_daily_fab_listing_function"

import pandas as pd
from os import listdir
from os.path import isfile, join
import sys
sys.path.append("C:\\Users\\cwilson\\Documents\\Python")
from read_daily_fab_listing_function import compile_daily_fab




daily_fab_python_folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
files = [f for f in listdir(daily_fab_python_folder) if isfile(join(daily_fab_python_folder,f))]


folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
file = "2020.09 Daily Fabrication.xlsx"



out = pd.DataFrame()



for file in files[-2:]:
    print(file)
    out = out.append(compile_daily_fab(folder, file))




fitters = out.Fitter.unique()
welders = out.Welder.unique()

out['unit wt'] = out['Wt'] / out['Qty']



out = out[out['unit wt'] > 100]
