# -*- coding: utf-8 -*-
"""
Created on Thu Mar 11 08:52:50 2021

@author: CWilson
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials


def init_google_sheet(key):
    
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    json_file = 'C:\\Users\\cwilson\\Documents\\Python\\read-daily-fab-c5ac66b791a1.json'
    
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)
    
    client = gspread.authorize(creds)
    
    # open the worksheet as a whole
    sh = client.open_by_key(key)
    
    return sh