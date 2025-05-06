# -*- coding: utf-8 -*-
"""
Created on Mon Apr 21 16:05:04 2025

@author: Netadmin
"""

import datetime
from pathlib import Path
import os
from fitter_welder_stats.Fitter_Welder_Stats_v2 import fitter_welder_stats_month
from fitter_welder_stats.Fitter_Welder_stats_PDF_report import pdf_report


file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

today = datetime.datetime.now()
month = today.month - 1
month_name = datetime.datetime(today.year, month, 1).strftime('%B')
if month == 12:
    year = today.year - 1
else:
    year = today.year
    
aggregate_data = fitter_welder_stats_month(month, year)
''' Handle sending messages out about the missing data'''
#%%



csv_file = aggregate_data['filepath']


output_dir = Path().home() / 'documents' / 'FitterWelderStatsPDFReports'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    
    


for state in ['MD','DE','TN']:
    output_file = f"FitterWelderStats-{state}-{str(month).zfill(2)}-{year}_{file_timestamp}.pdf"
    output_filepath = output_dir / output_file
    pdfreport = pdf_report(state, aggregate_data=aggregate_data, output_pdf=output_filepath)
    pdfreport.build_report()
