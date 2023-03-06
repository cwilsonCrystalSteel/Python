# -*- coding: utf-8 -*-
"""
Created on Mon Apr 12 13:52:15 2021

@author: CWilson
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials


def init_google_sheet(key):
    
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    json_file = 'C:\\Users\\cwilson\\Documents\\Python\\Attendance Project\\attendance-310518-02c301c1099c.json'
    
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)
    
    client = gspread.authorize(creds)
    
    # open the worksheet as a whole
    sh = client.open_by_key(key)
    
    return sh