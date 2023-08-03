# -*- coding: utf-8 -*-
"""
Created on Thu May 20 08:13:06 2021

@author: CWilson
"""

import pandas as pd
import datetime
import glob
import os
import numpy as np
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from navigate_EVA_folder_function import get_df_of_all_lots_files_information
import sys
sys.path.append('c://users//cwilson//documents//python//Attendance Project//')
from attendance_google_sheets_credentials_startup import init_google_sheet as init_google_sheet_production_worksheet
sys.path.append('C:\\Users\\cwilson\\documents\\python\\Speedo_Dashboard')
from Post_to_GoogleSheet import get_production_worksheet_job_hours


# shop = 'CSM'
# get the data from the fab listing google sheet
# fablisting_df = grab_google_sheet(shop + ' QC Form', "06/01/2021", "02/15/2022")
# get the df of eva destinations 
# eva_destinations_df = get_df_of_all_lots_files_information()

# fill = apply_model_hours2(fablisting_df, how='model', fill_missing_values=True, shop=shop, return_missing_job_lots=True)
# fill['missing job lots'].to_excel('c:\\users\\cwilson\\downloads\\' + shop + ' missing job lots.xlsx')
# fill['df'].to_excel('c:\\users\\cwilson\\downloads\\' + shop + ' fablisting with model filled in.xlsx')

def apply_model_hours2(fablisting_df, how='model', fill_missing_values=False, shop='min', return_missing_job_lots=False):
    
    missing_job_lots = pd.DataFrame(columns=['job','lot','reason', 'shops'])
    
    critical_columns = ['JOB NUMBER', 'SEQUENCE', 'PAGE', 'PRODUCTION CODE', 'QTY','SHAPE', 'LABOR CODE', 'MAIN MEMBER', 'TOTAL MANHOURS']

    
    # creat the lot name from the number
    fablisting_df['Lot Name'] = 'T' + fablisting_df['Lot #'].str.zfill(3)
    # 
    if how == 'model':
        # get the different job & lot combinations (and the count of # of records for each, but that doesn't matter)
        joblots = fablisting_df.groupby(['Job #','Lot Name']).size()
        
        
        if fablisting_df.shape[0]:
            df = pd.DataFrame()
            
            jobs = pd.unique(joblots.index.get_level_values(level=0))
            for job in jobs:
                print(job)
                
                chunk_job = fablisting_df[fablisting_df['Job #'] == job]
                # if for some reason the chunk returns no rows we have an issue!
                if not chunk_job.shape[0]:
                    print('No rows found for chunk_job: {}'.format(job))
                    chunk_job['Hours Per Piece'] = np.nan
                    df = df.append(chunk_job)     
                    continue 
                # try to get the job's eva database open
                try:
                    xls_main = pd.read_excel('C://downloads//' + str(job) + '.xlsx')                
                except Exception:
                    try:
                        xls_main = pd.read_excel('C://downloads//' + str(int(job)) + '.xlsx')  
                    except Exception:
                        print('coudld not open job "database": C://downloads//{}.xlsx'.format(job))
                        chunk_job['Hours Per Piece'] = np.nan
                        df = df.append(chunk_job)
                        continue 
                
                lots = joblots.xs(job, level=0).index
                
                for lot_name in lots:
                    print(job, lot_name)
                    # get the job 
                    # job = job_lot[0]
                    # get the lot name
                    # lot_name = job_lot[1]
                    # get the chunk with only that job & lot
                    chunk = chunk_job[chunk_job['Lot Name'] == lot_name]
                    if not chunk.shape[0]:
                        print('no rows found for chunk: {}'.format(lot_name))
                        chunk['Hours Per Piece'] = np.nan
                        df = df.append(chunk)
                        continue
                    chunk = chunk.copy()
                    
                    # start widdling down the big ass list of EVA files
                    # eva = eva_destinations_df[(eva_destinations_df['lot'] == lot_name) & (eva_destinations_df['job'] == job)]
                    '''
                    shops = list(eva['shop'])
                    # get only that location's eva file
                    eva = eva[eva['shop'] == shop]
                    # this will fail if there is not lot file available after looking for the LOT NAME -> JOB -> SHOP                
                    try:
                        # the iloc[-1] ensures getting the newest file
                        eva_destination = eva.iloc[-1]['destination']
                    except:
                        print('There is no file for {}'.format(job_lot))
                        if lot_name == 'T000':
                            reason = 'No LOT on fablisting'
                        else:
                            reason = 'Cannot find file in EVA folder'
                        missing_job_lots = missing_job_lots.append({'job':job, 'lot':lot_name, 'reason':reason,'shops':shops}, ignore_index=True)
                        chunk['Hours Per Piece'] = np.nan
                        df = df.append(chunk)
                        continue
                        
                        '''
                    # try to  open the EVA xls file
                    try:
                        # xls_lot = pd.read_excel(eva_destination, header=2, engine='xlrd', sheet_name='RAW DATA', usecols=critical_columns)
                        # xls_main = pd.read_excel('C://downloads//' + str(job) + '.xlsx')
                        if 'T' in lot_name:
                            xls_lot_from_main = xls_main[(xls_main['LOT'] == lot_name) | (xls_main['LOT'] == lot_name.replace('T',''))]
                        else:
                            xls_lot_from_main = xls_main[xls_main['LOT'] == lot_name]
                    except:
                        # print('Cannot open {}'.format(eva.iloc[-1]['basename']))
                        print('Cannot open {}'.format(lot_name))
                        ''' THIS IS WHERE I WOULD INFILL WHEN I CANNOT GET THE LOT EVA HOURS '''
                        # missing_job_lots = missing_job_lots.append({'job':job, 'lot':lot_name, 'reason':'Cannot open file: ' + eva_destination,'shops':shops}, ignore_index=True)
                        chunk['Hours Per Piece'] = np.nan
                        df = df.append(chunk)
                        continue
                    
                    # group by the page
                    xls_lot_paged = xls_lot_from_main.groupby('PAGE').sum()
                    # get only the main members and then get the sum of the QTY
                    xls_lot_mainmember_qty = xls_lot_from_main[xls_lot_from_main['MAIN MEMBER'] == 1].groupby('PAGE').sum()['QTY']
                    # drop the qty from the grouped df
                    xls_lot_paged = xls_lot_paged.drop(columns='QTY')
                    # add in the new calculated quantity
                    xls_lot_paged = xls_lot_paged.join(xls_lot_mainmember_qty)
                    # calcualte man hours per piece
                    xls_lot_paged['TOTAL MANHOURS PER PIECE'] = xls_lot_paged['TOTAL MANHOURS'] / xls_lot_paged['QTY']
                    
                    
                    
                    
                    ''' This is to get rid of the revision numbers on the pcmark os that I can join the manhours '''
                    # get a copy of the pcmarks column
                    pcmarks = chunk['Piece Mark - REV'].copy()
                    # split the piece marks based on a hyphen - this is incase they have the rev # next to it
                    pcmarks = pcmarks.str.split('-').str[0]
                    # get rid of any extra spaces in the piece mark 
                    pcmarks = pcmarks.str.strip()
                    # create a copy of chunk to prevent setwithcopy warning
                    chunk_copy = chunk.copy()
                    # set the copy of chunk's pcmark column to the new pcmarks series that has no hyphens now
                    chunk_copy['Piece Mark - REV'] = pcmarks
                    # set chunk to equal the copy  
                    chunk = chunk_copy
                    # get rid of the copy 
                    del chunk_copy
                    # grab the current index - for later so you can put the index back to normal
                    current_index = chunk.index
                     # set the index to be piecemark so that i can join easily
                    chunk = chunk.set_index('Piece Mark - REV', drop=False)
                    # get the hours per piece from the grouped xls df
                    chunk['Hours Per Piece'] = xls_lot_paged['TOTAL MANHOURS PER PIECE']
                    # set the chunk index back 
                    chunk = chunk.set_index(current_index)
                    
                    
                    # appends chunk to df, but now with 'Hours Per Piece' column
                    df = df.append(chunk)

        # if the fablisting_df is empty, then do this stuff
        else:
            # just let df be a copy of fablisting_df -> basically just to keep the columns
            df = fablisting_df.copy()
            # just set the column of 'Hours Per Piece' to nan
            df['Hours Per Piece'] = np.nan

        # calculate the 'Earned Hours' of the pieces based on the quantity in fablisting
        df['Earned Hours'] = df['Quantity'] * df['Hours Per Piece']
        # figure out which pieces did not have a match for EVA hours in the .xls files
        df['Has Model'] = ~df['Earned Hours'].isna()
        
        if fill_missing_values == True:
            df = fill_missing_model_earned_hours(df, shop)
    
    elif how == 'model but Justins dumb way of getting average hours':
        # get the different job & lot combinations (and the count of # of records for each, but that doesn't matter)
        joblots = fablisting_df.groupby(['Job #','Lot Name']).size()
        
        
        if fablisting_df.shape[0]:
            df = pd.DataFrame()
            
            jobs = pd.unique(joblots.index.get_level_values(level=0))
            try:
                jobs = jobs.astype(int)
            except:
                'print could not convert jobs to ints'
                
            for job in jobs:
                print(job)
                
                chunk_job = fablisting_df[fablisting_df['Job #'] == job]
                # if for some reason the chunk returns no rows we have an issue!
                if not chunk_job.shape[0]:
                    print('No rows found for chunk_job: {}'.format(job))
                    chunk_job['Hours Per Pound'] = np.nan
                    df = df.append(chunk_job)     
                    continue 
                # try to get the job's eva database open
                try:
                    xls_main = pd.read_excel('C://downloads//' + str(job) + '.xlsx')                
                except Exception:
                    try:
                        xls_main = pd.read_excel('C://downloads//' + str(int(job)) + '.xlsx')  
                    except Exception:
                        print('coudld not open job "database": C://downloads//{}.xlsx'.format(job))
                        chunk_job['Hours Per Pound'] = np.nan
                        df = df.append(chunk_job)
                        continue 
                
                # get the second index level - which is the lots
                lots = joblots.xs(job, level=0).index
                
                for lot_name in lots:
                    print(job, lot_name)

                    chunk = chunk_job[chunk_job['Lot Name'] == lot_name]
                    if not chunk.shape[0]:
                        print('no rows found for chunk: {}'.format(lot_name))
                        chunk['Hours Per Pound'] = np.nan
                        df = df.append(chunk)
                        continue
                    
                    chunk = chunk.copy()
                    
           
                    # try to  open the EVA xls file
                    try:
                        if 'T' in lot_name:
                            xls_lot_from_main = xls_main[(xls_main['LOT'] == lot_name) | (xls_main['LOT'] == lot_name.replace('T',''))]
                        else:
                            xls_lot_from_main = xls_main[xls_main['LOT'] == lot_name]
                    except:
                        # print('Cannot open {}'.format(eva.iloc[-1]['basename']))
                        print('Cannot open {}'.format(lot_name))
                        ''' THIS IS WHERE I WOULD INFILL WHEN I CANNOT GET THE LOT EVA HOURS '''
                        # missing_job_lots = missing_job_lots.append({'job':job, 'lot':lot_name, 'reason':'Cannot open file: ' + eva_destination,'shops':shops}, ignore_index=True)
                        chunk['Hours Per Pound'] = np.nan
                        df = df.append(chunk)
                        continue
                    
                    if xls_lot_from_main.shape[0]:
                        
                        avg_per_pound = xls_lot_from_main['TOTAL MANHOURS'].sum() / xls_lot_from_main['WEIGHT'].sum()
                                           
                        ''' This is to get rid of the revision numbers on the pcmark os that I can join the manhours '''
                        # get a copy of the pcmarks column
                        pcmarks = chunk['Piece Mark - REV'].copy()
                        # split the piece marks based on a hyphen - this is incase they have the rev # next to it
                        pcmarks = pcmarks.str.split('-').str[0]
                        # get rid of any extra spaces in the piece mark 
                        pcmarks = pcmarks.str.strip()
                        # create a copy of chunk to prevent setwithcopy warning
                        chunk_copy = chunk.copy()
                        # set the copy of chunk's pcmark column to the new pcmarks series that has no hyphens now
                        chunk_copy['Piece Mark - REV'] = pcmarks
                        # set chunk to equal the copy  
                        chunk = chunk_copy
                        # get rid of the copy 
                        del chunk_copy
                        # grab the current index - for later so you can put the index back to normal
                        current_index = chunk.index
                         # set the index to be piecemark so that i can join easily
                        chunk = chunk.set_index('Piece Mark - REV', drop=False)
                        # get the hours per piece from the grouped xls df
                        chunk['Hours Per Pound'] = avg_per_pound
                        # set the chunk index back 
                        chunk = chunk.set_index(current_index)
                        # appends chunk to df, but now with 'Hours Per Piece' column
                        df = df.append(chunk)
                    else:
                        ''' fill in missing values if fill_missing_values=True '''
                        
                        chunk['Hours Per Pound'] = np.nan
                        df = df.append(chunk)
                        continue

        # if the fablisting_df is empty, then do this stuff
        else:
            # just let df be a copy of fablisting_df -> basically just to keep the columns
            df = fablisting_df.copy()
            # just set the column of 'Hours Per Piece' to nan
            df['Hours Per Pound'] = np.nan

        # calculate the 'Earned Hours' of the pieces based on the quantity in fablisting
        df['Earned Hours'] = df['Weight'] * df['Hours Per Pound']
        # figure out which pieces did not have a match for EVA hours in the .xls files
        df['Has Model'] = ~df['Earned Hours'].isna()
        
        if fill_missing_values == True:
            # first try to infill usijng the LOTS LOG tab of production schedule - we will override any values already grabbed if possible
            # grab the lots log google sheet
            ll = get_LOTS_log_eva_hours()
            
            # create a copy of the df to work on for this exercise
            df2 = df.copy()
            # create the key to join to the LL with 
            df2['LOTS Name'] = df2['Job #'].astype(int).astype(str) + '-' + df2['Lot Name']
            # inner merge so that we only get records that match in the LL
            df2_plus_ll = pd.merge(df2.reset_index(), ll, left_on=['LOTS Name'], right_on=['LOTS Name']).set_index('index')
            print('We were able to match {} records with LOTS Log eva hours'.format(df2_plus_ll.shape[0]))
            # calculate the number of earned horus of each piece
            df2_plus_ll['Earned Hours'] = df2_plus_ll['Weight'] * df2_plus_ll['LOT EVA per lb']
            # make sure the has model tag is correct
            df2_plus_ll['Has Model'] = ~df2_plus_ll['Earned Hours'].isna()
            # change the value of hours per pound to match that from LL
            df2_plus_ll['Hours Per Pound'] = df2_plus_ll['LOT EVA per lb']
            # cut down the columns in df to match those of df
            df2_plus_ll = df2_plus_ll[list(df.columns)]     
            # override values in df with values in df2 - hopefully with the EVA hours from LL
            df.loc[df2_plus_ll.index] = df2_plus_ll
            
            # then try to backfilll with the old HPT way
            df = fill_missing_model_earned_hours(df, shop)
            
    
    
    elif how == 'old way':
        df = fill_missing_model_earned_hours(fablisting_df, shop)
    
    # sort the df back to original order
    df = df.sort_index()
    
    if return_missing_job_lots == False:
        
        # return the dataframe back out to the function caller
        return df
    
    else:
        return{'df':df, 'missing job lots':missing_job_lots}


def get_LOTS_log_eva_hours():
    _ProductionWorksheetGooglekey = "1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4"

    sh = init_google_sheet_production_worksheet(_ProductionWorksheetGooglekey)
    # get the values from the shipping schedule as a list of lists
    worksheet = sh.worksheet('LOTS Log').get_all_values()
    
    ll = pd.DataFrame(worksheet[3:], columns=worksheet[2])    

    new_cols = {}
    for col in ll.columns:
        new_col = col.replace('\n', ' ')
        new_cols[col] = new_col
        
    # replace columns with new columns w/o line breaks
    ll = ll.rename(columns=new_cols)    
    
    ll = ll[['Job','Fabrication Site','LOTS Name','Tonnage','TOTAL MHRS']]
    ll = ll.rename(columns={'TOTAL MHRS':'LOT EVA Hours'})
    ll['LOT EVA Hours'] = pd.to_numeric(ll['LOT EVA Hours'], errors='coerce')
    ll = ll[~ll['LOT EVA Hours'].isna()]
    ll['Tonnage'] = pd.to_numeric(ll['Tonnage'])
    ll = ll[~ll['Tonnage'].isna()]
    ll['LOT EVA per lb'] = ll['LOT EVA Hours'] / (ll['Tonnage'] * 2000)
    return ll
    
    
def apply_model_hours1(fablisting_df, how='model', fill_missing_values=False, shop='min'):
    
    if how == 'model':
        # get the unique jobs from the dataframe
        fablisting_jobs = pd.unique(fablisting_df['Job #'])
        # the current location of where I am putting the .xls files from fabsuite
        current_eva_folder = 'c://downloads/'
        # get a list of all of the .xls files in the folder
        list_of_xlsx = glob.glob(current_eva_folder + "*.xlsx")
        # if there is data in the fablisting df, proceed    
        if fablisting_df.shape[0]:
            # initialize an empty dataframe that will be appended to
            df = pd.DataFrame()
            # loop thru each job that is present in the fablisting dataframe
            for job in fablisting_jobs:
                # get only that portion of the job
                chunk = fablisting_df[fablisting_df['Job #'] == job]
                # grab the current index - for later so you can put the index back to normal
                current_index = chunk.index
                # get a copy of the pcmarks column
                pcmarks = chunk['Piece Mark - REV'].copy()
                # split the piece marks based on a hyphen - this is incase they have the rev # next to it
                pcmarks = pcmarks.str.split('-').str[0]
                # get rid of any extra spaces in the piece mark 
                pcmarks = pcmarks.str.strip()
                # create a copy of chunk to prevent setwithcopy warning
                chunk_copy = chunk.copy()
                # set the copy of chunk's pcmark column to the new pcmarks series that has no hyphens now
                chunk_copy['Piece Mark - REV'] = pcmarks
                # set chunk to equal the copy  
                chunk = chunk_copy
                # get rid of the copy 
                del chunk_copy
    
                # find the xls file that corresponds to the job - I named the xls files just the job number so this is easy as fuck
                job_xlsx_filenames = [s for s in list_of_xlsx if str(job) in s]
                # if there are .xls files with the job name, proceed
                if job_xlsx_filenames:
                    # get the newest file with the job number based on creation timestamp
                    job_xlsx_filename = max(job_xlsx_filenames, key=os.path.getctime)   
                    
                    # this is to test to make sure that the header is the right row
                    # while loop to open and read the first column to make sure the header is 'JOB NUMBER'
                    # save the row number ID so that when the whole file gets opened it can use the correct row #
                    header_num = 0
                    xls_main_test = pd.read_excel(job_xlsx_filename, header=header_num, nrows=5, usecols=[0])
                    while xls_main_test.columns[0] != 'JOB NUMBER':
                        header_num += 1
                        xlsx_main_test = pd.read_excel(job_xlsx_filename, header=header_num, nrows=5, usecols=[0])
                    
                    
                    
                    # read the .xls file
                    xlsx = pd.read_excel(job_xlsx_filename, header=header_num)
                    
           
                    # create a duplicator column which combines the PAGE, PRODUCTION CODE. SHAPE, LABRO CODE
                    xlsx['Duplicator'] = xlsx['PAGE'] + xlsx['PRODUCTION CODE'] + xlsx['SHAPE'] + xlsx['LABOR CODE']
                    # drop any duplicate rows based on the duplicator column
                    xlsx_drop = xlsx.drop_duplicates(subset=['Duplicator'])
                    # Get only the pieces that are marked as main members
                    mainmembers = xlsx[xlsx['MAIN MEMBER'] == 1]
                    # get the quantitys from the main members
                    mainmembersqty = mainmembers.groupby(['PAGE']).sum()['QTY']
                    # group the xls without duplicates by the PAGe & sum data
                    xlsx_drop_by_part = xlsx_drop.groupby(['PAGE']).sum()
                    # the quantity needs to be set the quantity of just the main members
                    xlsx_drop_by_part['QTY'] = mainmembersqty
                    # recalculate the man hours per piece
                    xlsx_drop_by_part['TOTAL MANHOURS PER PIECE'] = xlsx_drop_by_part['TOTAL MANHOURS'] / xlsx_drop_by_part['QTY']
                    # set the index to be piecemark so that i can join easily
                    chunk = chunk.set_index('Piece Mark - REV', drop=False)
                    # get the hours per piece from the grouped xls df
                    chunk['Hours Per Piece'] = xlsx_drop_by_part['TOTAL MANHOURS PER PIECE']
                    # set the chunk index back 
                    chunk = chunk.set_index(current_index)
                    
                    
                    
                # if there are no job .xls files, just set the 'Hours per piece' to be NAN
                else:
                    # setting the col to nan
                    chunk['Hours Per Piece'] = np.nan
    
                # appends chunk to df, but now with 'Hours Per Piece' column
                df = df.append(chunk)
        
        # if the fablisting_df is empty, then do this stuff
        else:
            # just let df be a copy of fablisting_df -> basically just to keep the columns
            df = fablisting_df.copy()
            # just set the column of 'Hours Per Piece' to nan
            df['Hours Per Piece'] = np.nan
            
        # calculate the 'Earned Hours' of the pieces based on the quantity in fablisting
        df['Earned Hours'] = df['Quantity'] * df['Hours Per Piece']
        # figure out which pieces did not have a match for EVA hours in the .xls files
        df['Has Model'] = ~df['Earned Hours'].isna()
        
        if fill_missing_values == True:
            df = fill_missing_model_earned_hours(df, shop)
    
    elif how == 'old way':
        df = fill_missing_model_earned_hours(fablisting_df, shop)
    
    # return the dataframe back out to the function caller
    return df


def load_averages_excel(shop):
    averages = pd.read_excel('C:\\downloads\\averages.xlsx')
    # sort the averages by the horus per ton
    averages = averages.sort_values(by='Hours per Ton')
    
    if shop == 'min':
        # keep the first of the duplicates
        averages = averages.drop_duplicates(subset='Job', keep='first')
    
    elif shop == 'max':
        # keep the last of the duplicates
        averages = averages.drop_duplicates(subset='Job', keep='last')       
    
    elif (shop == 'CSM') or (shop == 'FED') or (shop == 'CSF'):
        # get only the duplicated jobs
        duplicates = averages[averages['Job'].duplicated(keep=False)]
        # drop all of the duplicates from the averages df
        averages = averages.drop(index=duplicates.index)
        
        # go thru each job to see if the shop is present for that job
        for job in pd.unique(duplicates['Job']):
            # only get that job from the duplicates df
            chunk = duplicates[duplicates['Job'] == job]
            # only get the ones for that shop
            chunk = chunk[chunk['Shop'] == shop]
            # if the chunk has zero rows, take the minimum value
            if chunk.shape[0] == 0:
                # grab the duplicates for that job again
                chunk = duplicates[duplicates['Job'] == job]
                # keep the smallest hours per ton
                chunk = chunk.drop_duplicates(subset='Job', keep='first')
            
            # put the chunk back into the averages df
            averages = averages.append(chunk)
            
    averages = averages.set_index('Job')
    
    return averages


def fill_missing_model_earned_hours(fablisting_df, shop):
    fablisting_df = fablisting_df.copy()
    # this is the file that has the 
    averages = load_averages_excel(shop)
    
    # change the name for simplicity sake
    df = fablisting_df
    
    
    
    
    if 'Has Model' in df.columns:
        # get the pieces without model hours
        try:
            production_worksheet_hpt = get_production_worksheet_job_hours()
            
            no_model_search_jobs_shops = df[~df['Has Model'] & df['Earned Hours'].isna()]
            production_worksheet_hpt_shop = production_worksheet_hpt[production_worksheet_hpt['Shop'] == shop]
            
            nada_1 = pd.merge(no_model_search_jobs_shops, production_worksheet_hpt_shop, on='Job #', how='left')
            nada_1 = nada_1.set_index(no_model_search_jobs_shops.index)
            nada_1['Earned Hours'] = nada_1['HPT'] * nada_1['Weight']/2000
            nada_1['Hours Per Pound'] = nada_1['HPT']
            no_model_search_jobs_shops.loc[nada_1.index,:] = nada_1        
            # update the df with the newfound HPT values
            df.loc[no_model_search_jobs_shops.index] = no_model_search_jobs_shops
            
            # now try to get matches based on only the Job 
            no_model_search_just_jobs = df[~df['Has Model'] & df['Earned Hours'].isna()]
            nada_2 = pd.merge(no_model_search_just_jobs, production_worksheet_hpt, on='Job #', how='left')
            nada_2 = nada_2.set_index(no_model_search_just_jobs.index)
            nada_2['Earned Hours'] = nada_2['HPT'] * nada_1['Weight'] / 2000
            nada_2['Hours Per Pound'] = nada_2['HPT']
            no_model_search_just_jobs.loc[nada_2.index,:] = nada_2            
            #update the df with the newfound HPT values
            df.loc[no_model_search_just_jobs.index] = no_model_search_just_jobs
        except:
            print('Get_model_estimate_hours_attached_to_fablisting.py could not reach the google sheet for fill_missing_model_earned_hours')
        
        
        ''' Now we go back to try and fill in anything else from the Averages XLSX file '''
        no_model = df[~df['Has Model'] & df['Earned Hours'].isna()]
        # join the hours per ton based on the job #
        no_model = no_model.join(averages['Hours per Ton'], on='Job #')
        # calcualte the earned horus from the weight & average man hours per ton
        no_model['Earned Hours'] = no_model['Weight'] / 2000 * no_model['Hours per Ton']
        # drop the hours per ton rating
        # no_model = no_model.drop(columns=['Hours per Ton'])
        # set the no model rows back into the main dataframe 
        df.loc[no_model.index] = no_model
    
    # if there is no earned hours column -> has not been applied yet so just do old way of earned hours        
    else:
        # append the hours per ton
        no_model = df.join(averages['Hours per Ton'], on='Job #')
        # calculate the earned hours by tonnage
        no_model['Earned Hours'] = no_model['Weight'] / 2000 * no_model['Hours per Ton']
        # drop hours epr ton column
        # no_model = no_model.drop(columns = ['Hours per Ton'])
        # set df to be no_model
        df = no_model
        
    return df

