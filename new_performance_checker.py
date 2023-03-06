# -*- coding: utf-8 -*-
"""
Created on Mon Dec 14 15:18:27 2020

@author: CWilson
"""

"""
Goals:
    Create bar chart of fitter & welder performance over span of files
    
    Create tracker for tonnage and pieces by month
    
    Create histogram of weights of pieces
        Do so for shop over period of time & for each fitter & welder
        
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from os import listdir
from os.path import isfile, join
import sys
sys.path.append("C:\\Users\\cwilson\\Documents\\Python")
from read_daily_fab_listing_function import compile_daily_fab

DROP_BY_UNIT_WEIGHT = True
DROP_BY_QTY = True

DROP_BY_UNIT_WEIGHT_MIN_WEIGHT = 100
DROP_BY_QTY_AMOUNT = 10


daily_fab_python_folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
files = [f for f in listdir(daily_fab_python_folder) if isfile(join(daily_fab_python_folder,f))]


folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
file = "2020.09 Daily Fabrication.xlsx"



out = pd.DataFrame()



for file in files[-2:]:
    print(file)
    out = out.append(compile_daily_fab(folder, file), ignore_index=True)
out = out.reset_index(drop=True)



fitters = out.Fitter.unique()
welders = out.Welder.unique()

out['unit wt'] = out['Wt'] / out['Qty']


if DROP_BY_UNIT_WEIGHT:
    out = out[out['unit wt'] > DROP_BY_UNIT_WEIGHT_MIN_WEIGHT]
if DROP_BY_QTY:
    out = out[out['Qty'] < DROP_BY_QTY_AMOUNT]
    

def histogram_of_weights(num_bins, weight_series):
    
    fig, ax = plt.subplots()
    n, bins, patches = ax.hist(weight_series, num_bins)
    ax.set_xlabel("Weight")
    ax.set_ylabel("# of Occurences")
    ax.set_title("Histogram of Weight of Main Marks from: DATE TO DATE")
    plt.show()

# Shows only histogram below 2000 lbs    
# histogram_of_weights(50, out['unit wt'][out['unit wt'] <2000])
histogram_of_weights(50, out['unit wt'])




fitters_qty = pd.Series()
for fitter in fitters:
    fitters_qty[str(fitter)] = out['Qty'][out['Fitter'] == fitter].sum()

welders_qty = pd.Series()
for welder in welders:
    welders_qty[str(welder)] = out['Qty'][out['Welder'] == welder].sum()


num_days = (max(out['Date']) - min(out['Date'])).days

def order_by_sum(series, min_ave_pcs_completed=0, number_of_days=1):
    s = series.sort_values(ascending=False)
    s = s[s/number_of_days > min_ave_pcs_completed]
    
    return s

fitters_qty = order_by_sum(fitters_qty, min_ave_pcs_completed=.25, number_of_days=num_days)
welders_qty = order_by_sum(welders_qty, min_ave_pcs_completed=.25, number_of_days=num_days)






def barchart_of_production_by_person(series, title='', ylabel=''):
    x = range(len(series))
    y = series.values
    xlabels = series.index
    fig, ax = plt.subplots()
    ax.bar(x,y, align='center')
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.set_xticks(x)
    plt.xticks(rotation=90)
    ax.set_xticklabels(xlabels)
    plt.tight_layout()
    plt.show()

barchart_of_production_by_person(welders_qty, title='Number of pieces completed by welder')
barchart_of_production_by_person(fitters_qty, title='Number of pieces completed by fitter')













