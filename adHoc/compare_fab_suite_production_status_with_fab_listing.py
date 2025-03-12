# -*- coding: utf-8 -*-
"""
Created on Mon Mar 22 14:43:00 2021

@author: CWilson
"""

import glob
import os
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
import pandas as pd


folder = "C:\\Users\\cwilson\\Downloads\\5-18-2021\\*.csv"
output = "C:\\Users\\cwilson\\Downloads\\fab_differences\\"
files = glob.glob(folder)


this_year = grab_google_sheet("CSM QC Form", "01/01/2020", "12/31/2021")



for file in files:
    job = file[37:41]
    fab_suite = pd.read_csv(file)
    fab_suite = fab_suite[~fab_suite['Ship Status'].isna()]
    
    fab_listing = this_year[this_year['Job #'] == int(job)]
    
    # get the fabsuite pc mark list
    fs = fab_suite['Mark'].tolist()
    # get the fablisting pc mark list
    fl = fab_listing[fab_listing.columns[5]].tolist()
    # find ones in fabsuite not in fablisting
    pcs_not_in_fab_listing = pd.DataFrame(columns = fab_suite.columns)
    pcs_not_in_fab_suite = pd.DataFrame(columns = fab_suite.columns)
    
    for pc in fs:
        to_append = fab_suite[fab_suite['Mark'] == pc]
        if not(pc in fl):
            pcs_not_in_fab_listing = pcs_not_in_fab_listing.append(to_append)
    
    for pc1 in fl:
        to_append = fab_listing[fab_listing[fab_listing.columns[5]] == pc1]
        if not(pc in fs):
            pcs_not_in_fab_suite = pcs_not_in_fab_suite.append(to_append)
    
    
    cols = pcs_not_in_fab_listing.columns
    for col in cols[3:7]:
        pcs_not_in_fab_listing[col] = pcs_not_in_fab_listing[col].map(lambda x: x.rstrip('#'))
        pcs_not_in_fab_listing[col] = pcs_not_in_fab_listing[col].astype(float)

    
    # for pcmark in fab_suite['Mark']:
        
        
    #     a = fab_listing[fab_listing[fab_listing.columns[4]] == pcmark]
        
    #     if not a.empty:
    #         print(pcmark)
    #         df_to_append = fab_suite[fab_suite['Mark'] == pcmark]
    #         pcs_not_in_fab_listing = pcs_not_in_fab_listing.append(df_to_append, ignore_index=True)
            
            
    if not pcs_not_in_fab_listing.empty:
        print(job + " has pieces not in fablisting")
        pcs_not_in_fab_listing.to_csv(output + job + " in fabsuite not in fab listing.csv",)
    
    if not pcs_not_in_fab_suite.empty:
        print(job + " has pieces not in fabsuite")
        pcs_not_in_fab_suite.to_csv(output + job + "in fablisting not in fabsuite.csv")
        