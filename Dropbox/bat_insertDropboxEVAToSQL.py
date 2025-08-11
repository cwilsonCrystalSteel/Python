# -*- coding: utf-8 -*-
"""
Created on Sat Mar 29 15:44:07 2025

@author: Netadmin
"""


import os
from pathlib import Path
import glob
import pandas as pd
import numpy as np
import datetime
import re
from utils.insertErrorToSQL import insertError
from Dropbox.insertDropboxEVAToSQL import import_dropbox_eva_to_SQL, insert_evaDropbox_log, delete_filepath_for_redo
from Dropbox.pullDropboxEVAFromSQL import return_select_evadropbox_where_filename


source = 'bat_insertDropboxEVAToSQL'
insert_evaDropbox_log(description='Begin checking Dropbox EVA files', source=source)

critical_columns = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'PRODUCTION CODE', 'QTY',
                    'SHAPE', 'LABOR CODE', 'MAIN MEMBER', 'TOTAL MANHOURS', 'WEIGHT']
# New for Fit/Weld eva hours request 2025-08-11
extra_columns = ['300 - LAY OUT', '302 -  MARK / TAG', '350 - PUNCH', 
                 '400 -  FIT UP', '450 - BURN', '500 - WELD', 
                 '550 - FULL PEN WELD', '600 - CLEAN']
now = datetime.datetime.now()


def determine_directory_path():
    possible_dir = ['Y:/','X:/','\\\\192.168.50.9\\Dropbox_(CSF)']
    for ii in possible_dir:
        if os.path.exists(Path(ii)):
            base_dir = Path(ii) / 'production control' / 'EVA REPORTS FOR THE DAY'
            print(f'Using the drive: "{base_dir}"')
            break
        else:
            continue
    
    # we did not find a match at this point!
    if not os.path.exists(Path(ii)):
        insertError('EVADropbox - determine_directory_path', 'could not reach any directory!')
        raise Exception('Could not determine directory for Dropbox EVA files.')
    
    return base_dir







def check_file_existence_in_db(excel_file):
    redo = False
    exists = False
    # check to see if we have already recorded the xls file before
    
    # query
    existence_df = return_select_evadropbox_where_filename(str(excel_file))
    # if there is rows to the returned sql, then it exists
    if existence_df.shape[0]:
        exists = True
        
        # if it exists, then we need to check when it was inserted/updated and when the file was created/modified
        
        # check when the file was created
        file_creation_dt = datetime.datetime.fromtimestamp(os.path.getctime(excel_file))
        file_modified_dt = datetime.datetime.fromtimestamp(os.path.getmtime(excel_file))
        
        # convert to datetime
        existence_df['insertedat'] = pd.to_datetime(existence_df['insertedat'])
        # convert to datetime
        existence_df['updatedat'] = pd.to_datetime(existence_df['updatedat'])
        # somehow get the maximum value of these 
        # sql_time = existence_df[['insertedat','updatedat']].max().max()
        sql_time = existence_df[['insertedat']].max().max()
        
        # if the insertedat or updatedat is BEFORE the later of creation/modification, we will redo it
        if sql_time < max(file_creation_dt, file_modified_dt):
            # if the SQL time was before the later of creation/modified, then it needs to be redone
            redo = True
            
   
    
    return {'redo':redo, 'exists':exists}




base_dir = determine_directory_path()


def determine_folders_Methodically(days_back=30):
    
    
    dates = [now - datetime.timedelta(days=i) for i in range(0,days_back)]
    
    # this is how we get the xpected filepath
    # Example:'Y:/production control/EVA REPORTS FOR THE DAY/2025/02 - FEBRUARY/02142025'
    # Year: {i.year}                                                : 2025
    # Month: {str(i.month).zfill(2)} - {i.strftime('%B').upper()}   : 02 - FEBRUARY
    # Date: {i.strftime('%m%d%Y')}                                  : 02142025
    dates_folderpaths = [rf"{i.year}/{str(i.month).zfill(2)} - {i.strftime('%B').upper()}/{i.strftime('%m%d%Y')}" for i in dates]
    
    dates_folderpaths = [base_dir / i for i in dates_folderpaths]
    
    dates_folderpaths = [i for i in dates_folderpaths if os.path.exists(i)]
    
    return dates_folderpaths


def determine_folders_Glob(days_back=30):

    # get the dates we want to check
    dates = [now - datetime.timedelta(days=i) for i in range(0,days_back)]
    # convert those to the folder path names
    date_strings = [i.strftime('%m%d%Y') for i in dates]
    # get all of the folders in the direcotyr
    all_folders = glob.glob(str(base_dir /'*'/'*'/'*'))
    # find any of those that end with the values of date_strings
    dates_folderpaths = [i for i in all_folders if i.endswith(tuple(date_strings))]
    # convert to Path
    dates_folderpaths = [Path(i) for i in dates_folderpaths]
    
    return dates_folderpaths



def determine_excel_files(path):
    # get all files in the directory that are xls
    xls_files = glob.glob(str(path / '*.xls'))
    # get all files in the directory that are xlsx
    xlsx_files = glob.glob(str(path / '*.xlsx'))
    # combine the 2 lists
    excel_files = xls_files + xlsx_files
    # get any files that are PH?
    ph = [i for i in excel_files if 'PH' in i]
    if len(ph):
        print(ph)
    # remove those PH? files from the list of excel files
    excel_files = [Path(i) for i in excel_files if not i in ph]
    
    return excel_files


def massage_lot_name(lot):
    
        
    ''' Try to massage the lot name '''
    if lot[0].lower() == 't' and len(lot) == 4:
        rule = 'LOT 0'
        
        
    elif lot[0].lower() == 't' and len(lot) == 3:
        old_lot = lot
        lot = 'T' + lot[1:].zfill(3)
        rule = 'LOT 1'
        
    elif lot.isnumeric():
        old_lot = lot
        lot = 'T' + lot.zfill(3)
        rule = 'LOT 2'
        
    elif lot[0].lower() == 't' and lot[1:4].isnumeric():
        old_lot = lot
        lot = lot[:4]
        rule = 'LOT 3'
        
        
    # this could get TKT### or TK###
    elif lot[0:2].lower() == 'tk':
        old_lot = lot
        
        # convert TK into TKT
        if not lot.lower().startswith('tkt'):
            split_on_number = re.split(r'(\d+)', lot)
            split_on_number[0] = 'TKT'
            lot = ''.join(split_on_number)
            rule = 'TKT 1'
            
        
        # make sure the number part is 3 digits 
        split_on_number = re.split(r'(\d+)', lot)
        if len(split_on_number[1]) < 3:
            split_on_number[1] = split_on_number[1].zfill(3)
            old_lot = lot
            lot = ''.join(split_on_number)
            rule = 'TKT 2'
            
        # no changes were made
        if lot == old_lot:
            rule = 'TKT 0'

        
    else:
        print(f'This lot does not fill other criteria\n{lot}\t{excel_file}')
        insertError(name='EVADropbox', description = f'Could not find a valid mapping of the Lot Name: {excel_file}')
        # write_to_logfile(day_dt, f'{lot} is does not match expected criterias: {excel_file}')
        rule = 'e'
        
    if rule.endswith('0'):
        pass
        # print(f'Rule {rule}: {lot}')
    elif rule == 'e':
        pass
    else:
        print(f'Rule {rule}: {old_lot} to {lot}')
        
    return lot, rule
     


def adhoc_get_all_excel_files():
    # i am using this function to try and get a better idea of all the rules we need to define
    all_folders = glob.glob(str(base_dir /'*'/'*'/'*'))
    all_folders = [Path(i) for i in all_folders]
    all_excel_lots = [file for folder in all_folders for file in determine_excel_files(folder)]

    basenames = [os.path.basename(i) for i in all_excel_lots]

    investigation = {}

    for i in basenames:
        lot = i.split('-')[1]
        
        lot, rule = massage_lot_name(lot)
        
        if not rule in investigation.keys():
            investigation[rule] = []
            
        investigation[rule].append(lot)
        
        
#%%        

folderpaths = determine_folders_Glob(days_back=30)
excel_files = [file for folder in folderpaths for file in determine_excel_files(folder)]

#%%

for excel_file in excel_files:
    
    
    file_work = check_file_existence_in_db(excel_file)
    
    # get the file's name
    basename = os.path.basename(excel_file)
    # split it on the hyphen
    basename_components = basename.split('-')
    
    # check if it needs redoing ---> delete & do re-insert
    if file_work['redo']:
        print(f'Attempting to delete file: {excel_file}')
    
        delete_filepath_for_redo(filepath_str = excel_file)
      
        #??? for now, 2025-03-25, i will just do an error entry
        # insertError(name='EVADropbox - redo = True', description = f'This file needs to be deleted and reinserted, or updated via the merge proc(?): {excel_file}')
        
      
    # continue
    elif file_work['exists']:
        print(f"Exists in database & skipping: {excel_file}")
        continue
      
    
    
    # THERE SHOULD ALWAYS BE 3 pIECS TO THE FILE NAME:
        # JOB
        # LOT
        # SHOP / SITE
    if len(basename_components) != 3:
        # try to find another way to map it?
        try:
            basename_components = re.split(r'-|\ ', basename)
            job_str = basename_components[0]
            lot = basename_components[1]
            lot_orig = lot.copy()
            shop = basename_components[2][:-4]
            print(f'Able to resolve the invalid filename to its parts:\n\t{job_str}-{lot}-{shop}')
        
        # error when we cant map it, and skip
        except:
            insertError(name='EVADropbox - len(basename_compnents) != 3', description = f'Did not find 3 parts to the file name: {excel_file}')
            continue
    # when it is 3 parts like we expect, get its 3 parts  
    else:
        # job # is the start of the filename
        job_str = basename_components[0]
        job_str = re.sub(r'[A-Z]', '', job_str.upper())
        # lot is the middle portion of the filename
        lot = basename_components[1]
        lot_orig = lot
        # chop off the '.xls' and only maintain the shop 
        shop = basename_components[2].split('.')[0]
        
        
    
    # we are going to iterate thru the sheet_names as numbers to find the right one b/c calling by name results in a lot of failures
    sheet_num = 0
    max_sheet_num = 6
    while sheet_num <= max_sheet_num:
        try:
            # open the current lots file -> fingers crossed the header is alwasy row 2
            # only maintain those 9 columns that we actually need to pass along -> smaller file sizes
            excel_lot = pd.read_excel(excel_file, header=2, engine='xlrd', sheet_name=sheet_num, usecols=critical_columns + extra_columns)
            # if for some reason it pulls it but the dataframe is shape (0,0)
            if not excel_lot.size:
                raise Exception
                
            break
        except:
            pass
        
        sheet_num += 1
    
    # sheet_num = 0
    # max_sheet_num = 6
    # # this will first try and get a file with the extra_columns, but will default to just the critical_columns if it cant get the extra_columns
    # while sheet_num <= max_sheet_num:
    #     try:
    #         # First attempt: critical + extra columns
    #         try:
    #             cols_to_use = critical_columns + extra_columns
    #             excel_lot = pd.read_excel(
    #                 excel_file,
    #                 header=2,
    #                 engine='xlrd',
    #                 sheet_name=sheet_num,
    #                 usecols=cols_to_use
    #             )
    #             if not excel_lot.size:
    #                 raise ValueError("Empty dataframe with extra columns")
    #         except Exception:
    #             # Fallback: just critical columns
    #             excel_lot = pd.read_excel(
    #                 excel_file,
    #                 header=2,
    #                 engine='xlrd',
    #                 sheet_name=sheet_num,
    #                 usecols=critical_columns
    #             )
    #             if not excel_lot.size:
    #                 raise ValueError("Empty dataframe with critical columns")
    
    #         break  # Success, exit the loop
    
    #     except Exception:
    #         sheet_num += 1
    
        
    # when we max out the sheet_name iterator - skip this LOT
    if sheet_num  == max_sheet_num:
        
        #write to error sql table
        insertError(name='EVADropbox', description = f'Could not find a valid sheet in the Excel File with LOT data: {excel_file}')
        continue     



    ''' 2025-04-04
    # we are going to do the grouping to get the worth of main members in python
    # it was too much to do it on-demand in sql
    '''
    
    # This is counting how many times Production Code == Page
    # when Production Code == Page, it is a main member 
    # Production Code can have minor marks and main members 
    # but the minor marks have hours and weights that need added into the main member 
    excel_lot['matched_quantity'] = np.where(excel_lot['PRODUCTION CODE'] == excel_lot['PAGE'], excel_lot['QTY'], 0.0)
    
    # now calcualte the fit hours & weld hours
    excel_lot['fit_hours'] = excel_lot[['300 - LAY OUT','302 -  MARK / TAG','350 - PUNCH','400 -  FIT UP']].sum(axis=1)
    excel_lot['weld_hours'] = excel_lot[['500 - WELD','550 - FULL PEN WELD','600 - CLEAN']].sum(axis=1)
    
    # get the quantity, weight in pounds, and hours  
    main_members = excel_lot.groupby(['PAGE','SEQUENCE']).agg(
        quantity=('matched_quantity', 'sum'),
        pounds=('WEIGHT', 'sum'),
        hours=('TOTAL MANHOURS', 'sum'),
        fit_hours=('fit_hours', 'sum'),
        weld_hours=('weld_hours', 'sum')
    )
    
    # Add eva per quantity as a derived column
    main_members['evaperquantity'] = main_members['hours'] / main_members['quantity'].replace(0, np.nan)
    main_members['evaperquantity_fit'] = main_members['fit_hours'] / main_members['quantity'].replace(0, np.nan)
    main_members['evaperquantity_weld'] = main_members['weld_hours'] / main_members['quantity'].replace(0, np.nan)
    # move the PAGE & sequence back into a column
    main_members = main_members.reset_index()
    # make sure this goes into the db as as string
    main_members['PAGE'] = main_members['PAGE'].astype(str)
    # rename the 2 all caps to lower
    main_members = main_members.rename(columns={'PAGE':'page','SEQUENCE':'sequence'})
    
    # clean up the lot name 
    lot, rule = massage_lot_name(lot)     
    # set values of the lot/job into the dataframe 
    main_members['lotcleaned'] = lot
    # add these to go to the DB for tracability
    main_members['rawjob'] = int(job_str)
    main_members['rawlot'] = lot_orig
    main_members['rawshop'] = shop
    main_members['jobinfile'] = excel_lot.loc[0,'JOB NUMBER']
    
    # we need to get parts from the directory of the file
    directory, filename = os.path.split(excel_file)
    directory, folder_date = os.path.split(directory)
    directory, folder_month = os.path.split(directory)
    directory, folder_year = os.path.split(directory)
    
    # now add in the details of the file
    main_members['filepath'] = str(excel_file)
    main_members['filename'] = filename
    main_members['folderday'] = folder_date
    main_members['foldermonth'] = folder_month 
    main_members['folderyear'] = folder_year
    
    
        
    # insert into live table & do the merge proc
    try:
        import_dropbox_eva_to_SQL(main_members, source)
    except:
        insertError(name='EVADropbox - import_dropbox_eva_to_SQL', description = f'Could not find a valid sheet in the Excel File with LOT data: {excel_file}')
        continue                
    
                
                

insert_evaDropbox_log(description='Finished checking Dropbox EVA files', source=source)