# -*- coding: utf-8 -*-
"""
Created on Fri Jan 13 10:28:58 2023

@author: CWilson
"""

import pandas as pd

all_both_filepath = 'C:\\Users\\cwilson\\Documents\\Fitter_Welder_Performance_CSVs\\all_both_6month_July012022_December312022.csv'
all_both = pd.read_csv(all_both_filepath)
all_both = all_both[all_both['Quantity'] > 10]


hours_worked_folder = 'C:\\Users\\cwilson\\Documents\\Productive_Employees_Hours_Worked_Report\\week_by_week_hours_of_employees DE_formatted.xlsx'
df = pd.read_excel(hours_worked_folder)
