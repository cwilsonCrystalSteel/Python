# -*- coding: utf-8 -*-
"""
Created on Fri Mar 28 17:09:44 2025

@author: Netadmin
"""

from utils.attendance_google_sheets_credentials_startup import init_google_sheet as init_google_sheet_production_worksheet
import pandas as pd

_ProductionWorksheetGooglekey = "1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4"
sh = init_google_sheet_production_worksheet(_ProductionWorksheetGooglekey)
col_row = 2

def fix_cols(columns_list):
    # init a dict
    new_cols_dict = {}
    # iterate thru each column
    for col in columns_list:
        # replace any newlines with a space
        new_col = col.replace('\n', ' ')
        # make a dict entry of key = old, value = new
        new_cols_dict[col] = new_col
        
    # return dict to do pandas rename
    return new_cols_dict

def append_hyperlinks(ll):
    sheet_metadata = sh.fetch_sheet_metadata()
    worksheet_gid = None
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == 'LOTS Log':
            worksheet_gid = sheet['properties']['sheetId']
            break
    
    
    url_start = 'https://docs.google.com/spreadsheets/d/' + _ProductionWorksheetGooglekey
    url_end = f"/edit#gid={worksheet_gid}&range=A"
    
    ll['GoogleSheetLink'] = url_start + url_end + (ll.index + col_row + 1 + 1).astype(str)
    return ll


def pullEVALotsLogFromGoogle(drop_nonnumeric=True):

    
    # get the values from the shipping schedule as a list of lists
    worksheet = sh.worksheet('LOTS Log').get_all_values()
    
    ll = pd.DataFrame(worksheet[col_row + 1:], columns=worksheet[col_row])    
    # get rid of line breaks in the column names
    ll = ll.rename(columns=fix_cols(ll.columns))

    # only keep what we need
    ll = ll[['Job','Seq. #', 'Fabrication Site','LOTS Name','Tonnage','TOTAL MHRS']]
    ll = append_hyperlinks(ll)
    
    # rename the earned hours
    ll = ll.rename(columns={'TOTAL MHRS':'LOT EVA Hours'})
    # force conversion to a number
    ll['LOT EVA Hours'] = pd.to_numeric(ll['LOT EVA Hours'], errors='coerce')
    # make tonnage a number
    ll['Tonnage'] = pd.to_numeric(ll['Tonnage'], errors='coerce')
    if drop_nonnumeric:
        # get rid of anything that doesnt work to be forced into a number
        ll = ll[~ll['LOT EVA Hours'].isna()]
        # drop any unavailable tonnages 
        ll = ll[~ll['Tonnage'].isna()]
        
    # calcualte poundage of LOT horus
    ll['LOT EVA per lb'] = ll['LOT EVA Hours'] / (ll['Tonnage'] * 2000)
    return ll

    
    
