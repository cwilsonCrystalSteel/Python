# -*- coding: utf-8 -*-
"""
Created on Sun Apr  6 10:45:53 2025

@author: Netadmin
"""


from insertFablistingToSQL import insert_fablisting_to_live
from pullFablistingWithEVAFromSQL import get_fablisting_from_vreturnevafromfablisting

source = 'get_model_estimate_hours_attached_to_fablisting_SQL'

def call_to_insert(fablisting_df, sheet = None, source=source):
    if sheet is None:
        sheet = ''
    
    if not 'sheetname' in fablisting_df.columns:
        fablisting_df['sheetname'] = sheet
        
    if not 'hyperlink' in fablisting_df.columns:
        fablisting_df['hyperlink'] = ''
    
    insert_fablisting_to_live(fablisting_df, source=source)


def apply_model_hours_SQL(how='best', keep_diagnostic_cols=False):
    
    HOW_VALID_LIST = ['best','best_eva','best_hpt','eva_pcmark_dropbox','eva_lot_ave_lotslog',
                      'eva_job_ave_lotslog','eva_job_ave_dropbox','hpt_job_shop','hpt_job_ave']
    
    if not how in HOW_VALID_LIST:
        raise Exception(f'Value of how="{how}" did not match any values of {HOW_VALID_LIST}')
    
    
    fl_eva = get_fablisting_from_vreturnevafromfablisting()
    
    # put the names back to how they were before being inserted into live.fablisting
    rename_mapper = {'timestamp':'Timestamp', 
                     'job #':'Job #', 
                     'lot #':'Lot #', 
                     'quantity':'Quantity', 
                     'piece mark - rev':'Piece Mark - REV',
                     'weight':'Weight', 
                     'fitter':'Fitter', 
                     'fit qc':'Fit QC', 
                     'welder':'Welder', 
                     'weld qc':'Weld QC',
                     'does this piece have a defect?':'Does this piece have a defect?'}
    # do the renaming
    fl_eva = fl_eva.rename(columns=rename_mapper)
    
    
    if isinstance(how, list) and len(how) == 1:
        how = how[0]
        
    # if we are passed a list of the coalesce(value1, value2, ...)
    if isinstance(how, list):
    
        
        cols = []
        for i in how:
            if i in fl_eva.columns:
                cols.append(i)
            else:
                print(f"The provided column {i} was not found in the SQL Fablisting EVA view")
                
        
        fl_eva['Earned Hours'] = fl_eva[cols].bfill(axis=1).iloc[:, 0]


    
    elif how=='best':
        # essentially this is:
            # eva_pcmark_dropbox -> eva_lot_ave_lotslog -> eva_job_ave_lotslog ->
            # eva_job_ave_dropbox -> hpt_job_shop -> hpt_job_ave
        
        # get the eva coalesce and then the hpt coalesce
        fl_eva['Earned Hours'] = fl_eva['eva_hours'].combine_first(fl_eva['hpt_hours'])
    elif how == 'best_eva':
        fl_eva['Earned Hours'] = fl_eva['eva_hours']
    elif how == 'best_hpt':
        fl_eva['Earned Hours'] = fl_eva['hpt_hours']
    elif how == 'eva_pcmark_dropbox':
        fl_eva['Earned Hours'] = fl_eva['eva_hours_dropbox']
    elif how == 'eva_lot_ave_lotslog':
        fl_eva['Earned Hours'] = fl_eva['eva_hours_lotslog']
    elif how == 'eva_job_ave_lotslog':
        fl_eva['Earned Hours'] = fl_eva['eva_hours_lotslogjobaverage']
    elif how == 'eva_job_ave_dropbox':
        fl_eva['Earned Hours'] = fl_eva['eva_hours_dropboxjobaverage']  
    elif how == 'hpt_job_shop':
        fl_eva['Earned Hours'] = fl_eva['hpt_hours_jobshop']
    elif how == 'hpt_job_ave':
        fl_eva['Earned Hours'] = fl_eva['hpt_hours_jobaverage']
    
    if not keep_diagnostic_cols:
        
        diagnostic_cols = ['sheetname', 'pcmark', 'rev',
        'lot_3_digit', 'lot_with_t_start', 'shop', 'lotcleaned', 'eva_hours',
        'hpt_hours', 'eva_hours_dropbox', 'eva_hours_lotslog',
        'eva_hours_lotslogjobaverage', 'eva_hours_dropboxjobaverage',
        'hpt_hours_jobshop', 'hpt_hours_jobaverage', 'edjoin', 'elljoin',
        'elljobjoin', 'edjobjoin', 'hptjobshopjoin', 'hptjobjoin',
        'evaperquantity_dropbox', 'evaperpound_lotslog',
        'evaperpound_lotslogjobaverage', 'evaperpound_dropboxjobaverage',
        'hpt_jobshop', 'hpt_jobaverage', 'evasource']
        
        fl_eva = fl_eva.drop(columns=diagnostic_cols)
        
    
    
    
    return fl_eva