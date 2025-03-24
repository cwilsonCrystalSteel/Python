# -*- coding: utf-8 -*-
"""
Created on Mon Jul 26 08:11:44 2021

@author: CWilson
"""

from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2, fill_missing_model_earned_hours
# from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
import datetime
import pandas as pd
import sys
from TimeClock.pullGroupHoursFromSQL import get_date_range_timesdf_controller
from TimeClock.functions_TimeclockForSpeedoDashboard import return_information_on_clock_data
import os 
from pathlib import Path
# start_date = "07/01/2021"
# end_date = "07/31/2021"



yesterday = datetime.datetime.now() + datetime.timedelta(days=-1)
start_date = yesterday.strftime('%m/%d/%Y')
path = Path.home() / 'documents' / 'MDI' / 'Automatic'
if not os.path.exists(path):
    os.makedirs(path)
states = ['TN','DE','MD']

# basis = get_information_for_clock_based_email_reports(start_date, start_date)

def shape_check_before_to_excel(df_or_series, writer, sheet_name, indexTF):
    if df_or_series.shape[0]:
        df_or_series.to_excel(writer, sheet_name=sheet_name, index=indexTF)
    else:
        print('cannot send this df/series to excel b/c shape[0] == 0: {}'.format(sheet_name))

def do_mdi(basis=None, state='TN', start_date='01/01/2021', proof=True):
    '''
    basis: big dict that lists out absent df, Employee information df, direct hours df, indirect df, and raw clock information
    State: self explanatory
    Start date: requires format MM/DD/YYYY
    proof: spits out an excel file with all of the MDI data in seperate sheets
    '''
    
    # error handling -> can call the do_mdi without basis already generated
    if basis == None:
        # basis = get_information_for_clock_based_email_reports(start_date, start_date)
        times_df = get_date_range_timesdf_controller(start_date, start_date)
        basis = return_information_on_clock_data(times_df)
    
    
    # get the employee information df & sest the index to the employee name
    ei = basis['Employee Information'].set_index('Name')
    # init a dict that holds each state -> deprecated b/c now the function requires the state to be passed to it
    # still in use but now it only gets filled with the current state's data
    state_dict = {}
    # deprecated the for loop -> uses the functions argument State now
    # for state in states:
    
    # define the Google sheets sheetname for each state
    # also apply special changes to TN b/c scott wants prefab & fab instead of fit/qeld/saw/inventory
    if state == 'TN':
        sheet = 'CSM QC Form'        
    elif state == 'DE':
        sheet = 'CSF QC Form'
    elif state == 'MD':
        sheet = 'FED QC Form'
        
    
        
    # accoutn for the day being 6 am to 6 am for fablisting
    start_dt = datetime.datetime.strptime(start_date, '%m/%d/%Y')
    # subtract 7 hours to get the production for 2nd shift that started the day before - Justin
    # originally was + 6 hours to start the day at 6 am and go to 6 am the next day - origal
    # start_dt += datetime.timedelta(hours = -7)
    start_dt += datetime.timedelta(hours = 6)
    # add one day to make the end time start + one day
    end_dt = start_dt + datetime.timedelta(days=1)    
    # format as the date string
    end_date = end_dt.strftime('%m/%d/%Y')
    # get fablisting for all of start_date and all of end_date
    fablisting = grab_google_sheet(sheet, start_date, end_date)
    # get dates between yesterday at 6 am and today at 6 am
    fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
    # in case there is no records!
    if not fablisting.shape[0]:
        print(f'There were no records for {state} between {start_dt} and {end_dt}')
        return None
        
    # get the model hours attached to fablisting
    with_model = apply_model_hours2(fablisting, fill_missing_values=True, shop=sheet[:3])
    # sum up the new earned hours
    earned_new = with_model['Earned Hours'].sum()
    # get the old way of claculating earned hours
    old_way = fill_missing_model_earned_hours(fablisting, shop=sheet[:3])
    # put the old earned hours column onto the with_model df
    with_model = with_model.join(old_way[['Hours per Ton','Earned Hours']], rsuffix=' (old)')
    # sort the with_model dataframe by the creation time not the job
    with_model = with_model.sort_values('Timestamp')
    # sum up the old earned hours
    earned_old = old_way['Earned Hours'].sum()
    # calculate the tonnage
    tonnage = with_model['Weight'].sum() / 2000
    
    quantity = with_model['Quantity'].sum()
    ''' '''
    pieces_missing_model = with_model[with_model['Has Model'] == False]
    
    pieces_missing_model = pieces_missing_model[['Job #','Lot #','Piece Mark - REV', 'Weight', 'Quantity']]
    
    pieces_missing_model = pieces_missing_model.reset_index(drop=True)
    ''' Here I calcualte the difference between EVA & HPT models '''
    # only get the pieces that have EVA model
    pieces_hours_difference = with_model[with_model['Has Model'] == True]
    # cut down the columns to what I need
    pieces_hours_difference = pieces_hours_difference[['Job #','Lot #','Piece Mark - REV', 'Weight', 'Quantity','Earned Hours','Earned Hours (old)']]
    # rename columns to more fitting name
    pieces_hours_difference = pieces_hours_difference.rename(columns={'Earned Hours':'EVA', 'Earned Hours (old)':'HPT', 'Piece Mark - REV':'Pcmark'})
    # Drop any duplicate entries - incase the same piece was completed multiple times by different fitters/welders
    pieces_hours_difference = pieces_hours_difference.drop_duplicates(subset=['Pcmark'], keep='first')
    # get the value per individual piece
    pieces_hours_difference['EVA'] = pieces_hours_difference['EVA'] / pieces_hours_difference['Quantity']
    # get the value per indidividal piece
    pieces_hours_difference['HPT'] = pieces_hours_difference['HPT'] / pieces_hours_difference['Quantity']
    # get rid of quantity fcolumn
    pieces_hours_difference = pieces_hours_difference.drop(columns='Quantity')
    # calculate the difference in horus 
    pieces_hours_difference['Hour Diff.'] = pieces_hours_difference['EVA'] - pieces_hours_difference['HPT']
    # calcualte the percentage difference in hours
    pieces_hours_difference['Percent Diff.'] = 100 * abs(pieces_hours_difference['Hour Diff.']) / pieces_hours_difference['EVA']
    # sort by the percentage of difference
    pieces_hours_difference =  pieces_hours_difference.sort_values(by = 'Percent Diff.', ascending=False)
       
    ''' '''
    direct_df = basis['Direct'].copy()
    
    direct_df = direct_df[direct_df['Location'] == state]
    
    direct_df = direct_df.join(ei['Department'], on='Name')
    
    if 'Time In' in direct_df.columns:
        direct_df = direct_df.drop(columns='Time In')
        
    if 'Time Out' in direct_df.columns:
        direct_df = direct_df.drop(columns='Time Out')
    
    direct_df_departments = direct_df.groupby(['Job #','Cost Code','Department']).sum()
    
    direct_df_departments = direct_df_departments.sort_values(['Job #','Hours'], ascending=[False, False])
    
    direct = direct_df['Hours'].sum()
    
    if direct:
        efficiency_new = earned_new / direct
        
        efficiency_old = earned_old / direct
    else:
        efficiency_new = 0
        efficiency_old = 0

    indirect_df = basis['Indirect']
    
    indirect_df = indirect_df[indirect_df['Location'] == state]
    
    indirect_df = indirect_df.join(ei['Department'], on='Name')
    
    indirect = indirect_df['Hours'].sum()
    
    state_dict[state] = {'Earned (Model)': earned_new,
                         'Earned (Old)': earned_old,
                         'Direct': direct,
                         'Efficiency (Model)': efficiency_new,
                         'Efficiency (Old)': efficiency_old,
                         'Indirect': indirect,
                         'Tons': tonnage,
                         '# Pcs': quantity,}
    
    state_series = pd.Series(state_dict[state])
    
    state_series = state_series.rename(start_date)
    
    file_date = start_date.replace('/','-')
    
    if proof == True:
        print(path / (state + ' MDI ' + file_date + '.xlsx'))
        with pd.ExcelWriter(path / (state + ' MDI ' + file_date + '.xlsx')) as writer:
            # state_series.to_excel(writer, 'MDI')
            # pieces_missing_model.to_excel(writer, 'Missing Model Pieces')
            # direct_df_departments.to_excel(writer, 'LOT Department Breakdown')
            # direct_df.to_excel(writer, 'Direct Hours', index=False)
            # indirect_df.to_excel(writer, 'Indirect Hours', index=False)
            # with_model.to_excel(writer, 'Fablisting', index=False)
            # pieces_hours_difference.to_excel(writer, 'EVA vs HPT', index=False)
            
            shape_check_before_to_excel(state_series, writer, sheet_name='MDI', indexTF=True)
            shape_check_before_to_excel(pieces_missing_model, writer, sheet_name='Missing Model Pieces', indexTF=True)
            shape_check_before_to_excel(direct_df_departments, writer, sheet_name='LOT Department Breakdown', indexTF=True)
            shape_check_before_to_excel(direct_df, writer, sheet_name='Direct Hours', indexTF=False)
            shape_check_before_to_excel(indirect_df, writer, sheet_name='Indirect Hours', indexTF=False)
            shape_check_before_to_excel(with_model, writer, sheet_name='Fablisting', indexTF=False)
            shape_check_before_to_excel(pieces_hours_difference, writer, sheet_name='EVA vs HPT', indexTF=False)
                
    
    return {'MDI Summary':state_series.to_frame(), 'Missing Pieces':pieces_missing_model, 'Direct by Department':direct_df_departments, 'EVA vs HPT':pieces_hours_difference}


def verify_mdi(state, start_date, end_date, proof=False):
    
    state_df = pd.DataFrame()
    
    start_dt = datetime.datetime.strptime(start_date,'%m/%d/%Y')
    end_dt = datetime.datetime.strptime(end_date,'%m/%d/%Y')


    for day in range(0,(end_dt - start_dt).days + 1):
        # get the current datetime 
        dt = start_dt + datetime.timedelta(days=day)
        # conver to the string format i need
        date = dt.strftime('%m/%d/%Y')
        # get that day's information
        # basis = get_information_for_clock_based_email_reports(date, date)
        times_df = get_date_range_timesdf_controller(date, date)
        basis = return_information_on_clock_data(times_df)

        # convert basis to MDI format
        this_days_mdi = do_mdi(basis, state, date, proof=False)['MDI Summary']
        
        this_days_mdi = this_days_mdi.squeeze()
        # set the name of the series to be the date
        # this_days_mdi = this_days_mdi.rename(date)
        #append the series to the dataframe
        state_df[this_days_mdi.name] = this_days_mdi
    
    state_df = state_df.transpose()
    state_df['Weight'] = state_df['Tons'] * 2000
    state_df = state_df.fillna(0)
    start_date = start_date.replace('/', '-')
    end_date = end_date.replace('/','-')
    file = path / (state + ' Verification ' + start_date + ' to ' + end_date +'.xlsx')
    with pd.ExcelWriter(file) as writer:
        # state_df.to_excel(writer, 'Verification')
        shape_check_before_to_excel(state_df, writer, sheet_name='Verification', indexTF=True)
    return state_df



def eva_vs_hpt(start_date, end_date, proof=True):
    start_dt = datetime.datetime.strptime(start_date, '%m/%d/%Y')
    end_dt = datetime.datetime.strptime(end_date, '%m/%d/%Y') 
    
    #timespan for file naming 
    timespan = (end_dt - start_dt).days
    # add 6 hours to make it 6 am on the start dt
    start_dt += datetime.timedelta(hours = 6)
    # add one day to make the end time start + one day
    end_dt += datetime.timedelta(days=1, hours=6)    
    
    sheets = ['CSM QC Form', 'CSF QC Form', 'FED QC Form']
    
    # combine all the shops 
    for sheet in sheets:
        print(sheet)
        # get fablisting for all of start_date and all of end_date
        fablisting = grab_google_sheet(sheet, start_date, end_date)
        # get dates between yesterday at 6 am and today at 6 am
        fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
        # when we have no records:
        if not fablisting.shape[0]:
            with_model  = None
            print(f'There were no records for {sheet} between {start_dt} and {end_dt}')
        else:
            # get the model hours attached to fablisting
            with_model = apply_model_hours2(fablisting, fill_missing_values=True, shop=sheet[:3])
            
            old_way = fill_missing_model_earned_hours(fablisting_df=fablisting, shop=sheet[:3])
            # put the old earned hours column onto the with_model df
            with_model = with_model.join(old_way[['Hours per Ton','Earned Hours']], rsuffix=' (old)')
            
            with_model = with_model[['Job #','Lot #','Quantity','Piece Mark - REV','Weight','Earned Hours','Earned Hours (old)','Has Model']]
            
            with_model = with_model.rename(columns={'Earned Hours':'EVA', 'Earned Hours (old)':'HPT', 'Piece Mark - REV':'Pcmark'})
            
            with_model['Shop'] = sheet[:3]
        
        # if all_fab does not exist, and we have with_models
        if not 'all_fab' in locals() and not with_model is None:
            all_fab = with_model.copy()
        # if all_fab exists, and we have with_models
        elif not with_model is None:
            all_fab = pd.concat([all_fab, with_model])
       
    # this is for when NONE of the shops had any fablisting records
    if not 'all_fab' in locals():
        return None
    
    try:
        all_fab['Weight'] = all_fab['Weight'].apply(pd.to_numeric, errors='coerce')
        all_fab['EVA'] = all_fab['EVA'].apply(pd.to_numeric, errors='coerce')
        all_fab['HPT'] = all_fab['HPT'].apply(pd.to_numeric, errors='coerce')
        all_fab['Quantity'] = all_fab['Quantity'].apply(pd.to_numeric, errors='coerce')
    except:
        print('For some damn reason one of the all_fab number columns wont convert to numeric')
    
    missing_pieces = all_fab[all_fab['Has Model'] == False]
    
    missing_pieces = missing_pieces[['Job #', 'Lot #','Pcmark','Quantity','Weight','Shop']]
    
    missing_pieces = missing_pieces.groupby(['Job #','Lot #','Pcmark','Shop']).sum()
    
    missing_pieces = missing_pieces.reset_index()
    
    missing_pieces = missing_pieces[['Job #','Lot #','Pcmark','Quantity','Weight','Shop']]
    
    missing_by_lot = missing_pieces.groupby(['Job #', 'Lot #','Shop']).agg({
        'Quantity': 'sum',  # Sum numerical columns
        'Weight': 'sum',
        'Pcmark': lambda x: ', '.join(x)  # Concatenate strings
    }).reset_index()
    
    missing_by_lot = missing_by_lot.reset_index()
    
    missing_by_lot = missing_by_lot.sort_values(by=['Weight','Quantity','Job #'], ascending=False)
    
    eva_vs_hpt = all_fab[all_fab['Has Model'] == True]
    
    eva_vs_hpt = eva_vs_hpt.drop(columns=['Has Model'])
    
    eva_vs_hpt = eva_vs_hpt.groupby(['Job #','Lot #','Pcmark']).sum().reset_index()
    
    eva_vs_hpt['Tons'] = eva_vs_hpt['Weight'] / 2000
    # eva_vs_hpt = eva_vs_hpt.drop_duplicates('Pcmark', keep='first')
    
    eva_vs_hpt = eva_vs_hpt[['Job #','Lot #','Pcmark','Tons','Quantity','EVA','HPT']]
    
    eva_vs_hpt['Hr. Diff'] = eva_vs_hpt['EVA'] - eva_vs_hpt['HPT']
    
    eva_vs_hpt['% Diff'] = (abs(eva_vs_hpt['Hr. Diff']) / eva_vs_hpt['EVA'])
    
    eva_vs_hpt = eva_vs_hpt.sort_values('% Diff', ascending=False)
    
    eva_vs_hpt_by_job = eva_vs_hpt.groupby('Job #').sum()
    
    eva_vs_hpt_by_job['Hr. Diff'] = eva_vs_hpt_by_job['EVA'] - eva_vs_hpt_by_job['HPT']
    
    eva_vs_hpt_by_job['% Diff'] = (abs(eva_vs_hpt_by_job['Hr. Diff']) / eva_vs_hpt_by_job['EVA'])
    
    eva_vs_hpt_by_job = eva_vs_hpt_by_job.sort_values(by='% Diff', ascending=False)
    
    eva_vs_hpt_by_job = eva_vs_hpt_by_job.reset_index()
    
    eva_vs_hpt_by_lot = eva_vs_hpt.groupby(['Job #','Lot #']).sum()
    
    eva_vs_hpt_by_lot['Hr. Diff'] = eva_vs_hpt_by_lot['EVA'] - eva_vs_hpt_by_lot['HPT']
    
    eva_vs_hpt_by_lot['% Diff'] = (abs(eva_vs_hpt_by_lot['Hr. Diff']) / eva_vs_hpt_by_lot['EVA'])
    
    eva_vs_hpt_by_lot = eva_vs_hpt_by_lot.sort_values(by='% Diff', ascending=False)
    
    eva_vs_hpt_by_lot = eva_vs_hpt_by_lot.reset_index()
    
    path = Path.home() / 'documents' / 'EVA_VS_HPT' / 'Automatic'
    if not os.path.exists(path):
        os.makedirs(path)

    file_date = end_date.replace('/','-')
    timespan_str = '_' + str(timespan) + 'days'
    
    filename = None
    
    if proof == True:
        filename = path / ('EVA_vs_HPT ' + file_date + timespan_str + '.xlsx')
        print(filename)
        with pd.ExcelWriter(filename) as writer:
            # missing_pieces.to_excel(writer, 'Missing Pieces', index=False)
            # eva_vs_hpt.to_excel(writer, 'Pcmark', index=False)
            # eva_vs_hpt_by_lot.to_excel(writer, 'Lot', index=False)
            # eva_vs_hpt_by_job.to_excel(writer, 'Job', index=False)
            
            shape_check_before_to_excel(missing_pieces, writer, sheet_name='Missing Pieces', indexTF=False)
            shape_check_before_to_excel(eva_vs_hpt, writer, sheet_name='Pcmark', indexTF=False)
            shape_check_before_to_excel(eva_vs_hpt_by_lot, writer, sheet_name='Lot', indexTF=False)
            shape_check_before_to_excel(eva_vs_hpt_by_job, writer, sheet_name='Job', indexTF=False)
            
    
    return {'Pcmark':eva_vs_hpt, 
            'Job':eva_vs_hpt_by_job, 
            'Lot':eva_vs_hpt_by_lot, 
            'Missing':missing_pieces, 
            'Missing Summary': missing_by_lot,
            'Filename':filename}
   
    