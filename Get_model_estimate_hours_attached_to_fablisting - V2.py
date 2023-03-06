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



# get the data from the fab listing google sheet
# fablisting_df = grab_google_sheet('FED QC Form', "06/02/2021", "06/02/2021")





def apply_model_hours(fablisting_df, how='model', fill_missing_values=False, shop='min'):
    
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
                    
                    
                    # group based on the PAGE column - basically says everything that goes into that pcmark
                    # sum up all the hours associated with all pcs associated with the main marks
                    # this also makes the index the pcmarks! ez-pz
                    # xls_by_part = xls.groupby('PAGE').sum()
                    # # set the index to be the pcmarks
                    # chunk = chunk.set_index('Piece Mark - REV', drop=False)                  
                    # # set the chunk column of 'Hours Per Piece' equal to the man hours based on pcmark
                    # # made sure to download the total man hours by the count of MAIN MEMBER
                    # # cannot sum by quantity b/c there my be 3 detail pieces and 1 main member
                    # # so then it would make the qty=4 when in reality the qty of main marks is 1
                    # chunk['Hours Per Piece'] = xls_by_part['TOTAL MANHOURS'] / xls_by_part['MAIN MEMBER']
                    # # rest the chunks index back to the old index
                    # chunk = chunk.set_index(current_index)
                    
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


def fill_missing_model_earned_hours(fablisting_df, shop):
    # this is the file that has the 
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
    
    
    # change the name for simplicity sake
    df = fablisting_df
    
    if 'Has Model' in df.columns:
        # get the pieces without model hours
        no_model = df[~df['Has Model']]
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





''' Depreciated version using incorrect csv files '''
'''
def apply_model_hours2(fablisting_df):
    
    # get the unique jobs from the dataframe
    fablisting_jobs = pd.unique(fablisting_df['Job #'])
    
    # the current folder with eva files - this is where I dump files from the .zip from the RDP
    current_eva_folder = 'C://users/cwilson/documents/EVA_Estimate_csvs/5-21-2021/5-21-2021/'
    # list out all csv files in this folder
    list_of_csvs = glob.glob(current_eva_folder + "*.csv")
    # # get just the file basenames
    # file_names = [os.path.basename(x) for x in list_of_csvs]
    # # figure out which job numbers are availale 
    # job_numbers = [x[26:30] for x in file_names]
    
    
    # if there is data in the fablisting_df, then proceed
    if fablisting_jobs.shape[0]:
        # create a new dataframe
        df = pd.DataFrame()
        # loop thru each job available in the fablisting dataframe
        for job in fablisting_jobs:
            # get only the data for that job from fablisting_df
            chunk = fablisting_df[fablisting_df['Job #'] == job]
            # get the current index to reapply it at the end
            current_index = chunk.index
            # convert the pcmarks to get rid of the '-0' shit
            pcmarks = chunk['Piece Mark - REV'].copy()
            # get rid of the revision number, which is after the dash
            pcmarks = pcmarks.str.split('-').str[0]
            # copy the chunk to prevent SetWithCopyWarning
            chunk_copy = chunk.copy()
            # Set the Piece Mark column to be the pcmark without the revision
            chunk_copy['Piece Mark - REV'] = pcmarks
            # reassign the chunk dataframe
            chunk = chunk_copy
            # delete the copy of the variable
            del chunk_copy
            # set the index to be the Piece Mark - to easily join the hours per piece data based on pcmark
            chunk = chunk.set_index('Piece Mark - REV', drop=False)
            # get the EVA csv file for that job
            job_csv_filename = [s for s in list_of_csvs if str(job) in s]
            # if the list has an item in it, then do this stuff
            if job_csv_filename:
                # get the first element (which should be the only element b/c only one file per job)
                job_csv_filename = job_csv_filename[0]
                # read the EVA csv file 
                csv = pd.read_csv(job_csv_filename)
                # calculate the Hours per piece based on quantity and total hours per pcmark
                csv['Hours Per Piece'] = csv['Man Hours'] / csv['Quantity']
                # set the index to be pcmark - the natural key
                csv = csv.set_index('Production Code', drop=False)
                # assign the EVA column to the job chunk df - works based on unique & matching pcmarks as key
                chunk['Hours Per Piece'] = csv['Hours Per Piece']
                # reset the chunks index
                chunk = chunk.set_index(current_index)    
            
            # if the job is not in the list_of_csvs, then just set that column to be np.nan
            else:
                # just set the values of 'Hours Per Piece' to be nan
                chunk['Hours Per Piece'] = np.nan
                # reset the chunk's index
                chunk = chunk.set_index(current_index)
                
            # append the chunk back onto the outputted df
            df = df.append(chunk)
            

        
    # if there is no data in the fablisting_df, just give it some empty columns
    else:
        # just keep using the fablisting_df if there is no data in the df anyways (to maintain the structure of the columns)
        df = fablisting_df.copy()
        # Basically just creates a column for the last command in this function
        df['Hours Per Piece'] = np.nan
        # df['Earned Hours'] = np.nan
        
    
    # calculate the earned hours of the new fablisting df based on hours per piece and quantity in the df
    df['Earned Hours'] = df['Quantity'] * df['Hours Per Piece']    
    
    pieces_without_eva = df[df['Earned Hours'].isna()]
    
    return df 
'''