# -*- coding: utf-8 -*-
"""
Created on Fri Mar 19 13:04:12 2021

@author: CWilson
"""

import gspread
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
import datetime










def totals(sheet_name):
    
    
    end = datetime.datetime.today()
    end = end.replace(hour=23, minute=59, second = 59, microsecond=0)
    year_start = datetime.date(end.year, 1, 1)
    month_start = datetime.datetime(end.year, end.month, 1, 0, 0, 0)
    
    
    idx = (end.weekday() + 1) % 7
    week_start = end - datetime.timedelta(idx)
    week_start = week_start.replace(hour=0, minute=0, second=0, microsecond=0)
    
    
    year_start_string = year_start.strftime('%m/%d/%Y')
    end_string = end.strftime('%m/%d/%Y')
    
    
    
    this_year = grab_google_sheet(sheet_name, year_start_string, end_string)
    this_month = this_year[this_year['Timestamp'] >= month_start]
    this_week = this_year[this_year['Timestamp'] >= week_start]
    
    
    
    year_sum = this_year.sum(axis=0)
    year_weight = year_sum['Weight']
    year_piece = year_sum['Quantity']
    
    
    month_sum = this_month.sum(axis=0)
    month_weight = month_sum['Weight']
    month_piece = month_sum['Quantity']
    
    
    week_sum = this_week.sum(axis=0)
    week_weight = week_sum['Weight']
    week_piece = week_sum['Quantity']
    
    output = [[year_weight, year_piece], 
              [month_weight, month_piece], 
              [week_weight, week_piece]]
    
    return output


