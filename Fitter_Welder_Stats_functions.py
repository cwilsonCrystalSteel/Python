# -*- coding: utf-8 -*-
"""
Created on Mon Jun  7 15:42:29 2021

@author: CWilson
"""


import sys
sys.path.append('c://users//cwilson//documents//python//Weekly Shop Hours Project//')
from TimeClock_Tools_Employee_Department import download_most_current_employee_department_csv
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Grab_Defect_Log_Google_Sheet_Data import grab_defect_log
from TimeClock_Group_Hours import download_group_hours
import gspread
from google_sheets_credentials_startup import init_google_sheet
from Read_Group_hours_HTML import output_group_hours_by_job_code
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2
import pandas as pd
import glob
import os
import numpy as np
import datetime

def get_employee_name_ID():
    while True:
        try:
            download_most_current_employee_department_csv()
            # Grab all HTML files in downloads
            list_of_csvs = glob.glob("C://Users//Cwilson//downloads//*.csv") # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            employe_info_csvs = [f for f in list_of_csvs if "Employee Information" in f]
            # Get the most recent file for that state
            latest_employee_departments = max(employe_info_csvs, key=os.path.getctime)
            # check to make sure the creation date is today
            file_date = datetime.datetime.fromtimestamp(os.path.getctime(latest_employee_departments)).date()
            # if the file date is not today, throw an error
            if file_date != datetime.datetime.today().date():
                2 + "2"            
            
            print('Newest employee department: ', latest_employee_departments)
            ei = pd.read_csv(latest_employee_departments)
            os.remove(latest_employee_departments)
            break
        except:
            pass
    return ei



def download_employee_group_hours(start_date, end_date):
    count = 0
    while count <= 5:
        try:
            count += 1
            # download the group hours with the daterange
            download_group_hours(start_date, end_date)
            # list all html files in the download folder
            list_of_htmls = glob.glob("C://Users//Cwilson//downloads//*.html")
            # only get html files if they contain 'Hours'
            hours_htmls = [f for f in list_of_htmls if 'Hours' in f]
            # get the newest one
            latest_hours = max(hours_htmls, key=os.path.getctime)
            # check to make sure the creation date is today
            file_date = datetime.datetime.fromtimestamp(os.path.getctime(latest_hours)).date()
            # if the file date is not today, throw an error
            if file_date != datetime.datetime.today().date():
                2 + "2"            
            
            break
        except:
            pass
    
    print('Fixing the HTML file to a dataframe ')
    # process the html file to output as df
    job_hours_df = output_group_hours_by_job_code(latest_hours)
    # delete the file from download folder
    os.remove(latest_hours)
    return job_hours_df






def get_employee_hours(all_df, direct_df, indirect_df):
    
    grouped_direct = direct_df[['Name','Hours']].groupby(by='Name').sum()
    grouped_indirect = indirect_df[['Name','Hours']].groupby(by='Name').sum()
        
    # create a new dataframe which gets its columns appended
    # append_df = pd.DataFrame(columns=['Direct Hours','Indirect Hours','Total Hours'], index=grouped_direct.index)
    # append_df['Direct Hours'] = grouped_direct
    # append_df['Indirect Hours'] = grouped_indirect
    # append_df = append_df.fillna(0)
    # append_df['Total Hours'] = append_df['Direct Hours'] + append_df['Indirect Hours']

    
    grouped_direct = grouped_direct.rename(axis=1, mapper={'Hours':'Direct Hours'})
    grouped_indirect = grouped_indirect.rename(axis=1, mapper={'Hours':'Indirect Hours'})
    joining_df = grouped_direct[['Direct Hours']].join(grouped_indirect[['Indirect Hours']], how='outer')
    joining_df = joining_df.fillna(0)
    joining_df['Total Hours'] = joining_df['Direct Hours'] + joining_df['Indirect Hours']
    
    # set the index of the input df to the employee names
    all_df = all_df.set_index('Name', drop=False)
    # join the columns of the append_df to the input df
    all_df = all_df.join(joining_df)
    # reset the index back to numbers
    all_df = all_df.reset_index(drop=True)
    # calculate Percentage Direct of Total
    all_df['Direct/Total'] = all_df['Direct Hours'] / all_df['Total Hours']
    # calculate DL Efficiency
    all_df['DL Efficiency'] = all_df['Earned Hours'] / all_df['Direct Hours']
    # calculate Total Effiency
    all_df['TTL Efficiency'] = all_df['Total Hours'] / all_df['Direct Hours']
    # calculate DL Hrs/Ton
    all_df['DL Hrs/Ton'] = all_df['Direct Hours'] / all_df['Tonnage']
    # calculate TTL Hrs/Ton
    all_df['TTL Hrs/Ton'] = all_df['Total Hours'] / all_df['Tonnage']
    
    all_df['DL Efficiency'] = all_df['DL Efficiency'].replace(np.inf, 0)
    all_df['TTL Efficiency'] = all_df['TTL Efficiency'].replace(np.inf, 0)
    

    return all_df



def unsplit_shared_pieces(df1, col_name, multiple_employee_split_key, index_divider):

    # get only pieces with multiple fitters
    shared_pieces = df1[df1[col_name].str.contains(multiple_employee_split_key)]
    # remove anything that has N/A
    shared_pieces = shared_pieces[~shared_pieces[col_name].str.contains("N/A")]
    # remove anything that has NA
    shared_pieces = shared_pieces[~shared_pieces[col_name].str.contains("NA")]
    
    # get the original indexes of the shared_pieces to be removed
    shared_pieces_original_idx = shared_pieces.index.tolist()
    # create a new dataframe which will hold the broken apart entries
    split_pieces = pd.DataFrame(columns=shared_pieces.columns)
    
    for row in shared_pieces.index:
        # grab that slice of the shared pieces dataframe
        this_piece = shared_pieces.loc[row].copy()
        # break apart the fitter by the splitting key
        mult_employees = this_piece.loc[col_name].split(multiple_employee_split_key[-1])
        # get the number of employees
        num_employees = len(mult_employees)
        # divide the quantity by the number of people working on it    
        this_piece['Quantity'] = this_piece['Quantity'] / num_employees
        # divide the weight by the number of people working on it    
        this_piece['Weight'] = this_piece['Weight'] / num_employees
        # this is just in case the model hours are not applied to the fablisting df
        if 'Hours Per Piece' in this_piece.index:
            # divide the Hours Per Piece by the number of people working on it
            this_piece['Hours Per Piece'] = this_piece['Hours Per Piece'] / num_employees
            # divide the Earned Hours by the number of people working on it
            this_piece['Earned Hours'] = this_piece['Earned Hours'] / num_employees
        elif 'Hours per Ton' in this_piece.index:
            this_piece['Earned Hours'] = this_piece['Earned Hours'] / num_employees
    
        # iterate through each employee on the piece
        for i,emp in enumerate(mult_employees):
            # create a copy of the df/series
            split_piece = this_piece.copy()
            
            # create the decimal value that will get added to the index to create a unique index
            index_decimal = (i + 1) / index_divider
            # add the decimal value to the current index (which is the name of the series)
            new_index = split_piece.name + index_decimal
            # round the index to 2 decimal places to eliminate wacky shit like 45.11999999997 instead of 45.12)
            new_index = round(new_index, 4)
            # rename the series so each one has a unique identification
            split_piece = split_piece.rename(new_index)
            # rename whoever the fitter is 
            split_piece[col_name] = emp
            # append it to a new dataframe
            # split_pieces = split_pieces.append(split_piece)
            # make it a df not a series
            split_piece = split_piece.to_frame()
            # if the df is long and not wide - transpose it 
            if split_piece.shape[0] > split_piece.shape[1]:
                split_piece = split_piece.transpose()
            # append the new row
            split_pieces = pd.concat([split_pieces, split_piece], axis=0, ignore_index=True)            
    
    
    
    # remove those shared pieces rows from the original dataframe
    df1 = df1.drop(index = shared_pieces_original_idx)
    # add in the split_pieces to replace the shared_pieces
    # df1 = df1.append(split_pieces)
    df1 = pd.concat([df1, split_pieces], axis=0, ignore_index=True)
    # puts the pieces back in place where they belong
    df1 = df1.sort_index()
    
    return df1


def clean_and_adjust_fab_listing_for_range(state, start_date, end_date, earned_hours='old way'):
    
    if state == 'TN':
        fab_google_sheet_name = 'CSM QC Form'
        multiple_employee_split_key = '/'
    elif state == 'MD':
        fab_google_sheet_name = 'FED QC Form'
        multiple_employee_split_key = '\.'
    elif state == 'DE':
        fab_google_sheet_name = 'CSF QC Form'
        multiple_employee_split_key = '\.'

    # grab the sheet data
    df = grab_google_sheet(fab_google_sheet_name, start_date, end_date)
    # get the EVA hours applied to the dataframe
    df = apply_model_hours2(df, how=earned_hours, fill_missing_values=True, shop=fab_google_sheet_name[:3])
    
    # fix the split up pieces for multiple fitters -> column 6
    df = unsplit_shared_pieces(df, col_name="Fitter", multiple_employee_split_key=multiple_employee_split_key, index_divider=10)
    # fix the split up pieces for multiple welders -> column 8
    df = unsplit_shared_pieces(df, col_name="Welder", multiple_employee_split_key=multiple_employee_split_key, index_divider=100)
    # convert Job #, Fitter, Fitter QC, Welder, Welder QC to numbers
    df["Job #"] = df["Job #"].apply(pd.to_numeric, errors='coerce')
    df["Fitter"] = df["Fitter"].apply(pd.to_numeric, errors='coerce')
    df["Welder"] = df["Welder"].apply(pd.to_numeric, errors='coerce')
    # df["Fit QC"] = df["Fit QC"].apply(pd.to_numeric, errors='coerce')
    # df["Weld QC"] = df["Weld QC"].apply(pd.to_numeric, errors='coerce')
    
    # remove row if the fitter columns is NaN
    df = df[~df["Fitter"].isna()]
    # remove row if the Job # column (col 2) is NaN
    df = df[~df["Job #"].isna()]
    # get the unique fitters
    fitters = pd.unique(df["Fitter"])
    # get the unique fitter qc
    # fitters_qc = pd.unique(df["Fit QC"])
    # get the unique welders
    welders = pd.unique(df["Welder"])
    # get the unique welders qc
    # welders_qc = pd.unique(df["Weld QC"])
    # drop the null values from the array of fitters
    fitters = fitters[~pd.isnull(fitters)]
    # drop the null values from the array of fitters qc
    # fitters_qc = fitters_qc[~pd.isnull(fitters_qc)]
    # drop the null values from the array of welders
    welders = welders[~pd.isnull(welders)]
    # drop the null values from the array of welders qc
    # welders_qc = welders_qc[~pd.isnull(welders_qc)]
    # convert the defect column to just lowercase
    # df[df.columns[-3]] = df[df.columns[-3]].str.lower()
        
    return {'Fab df': df, 'Fitter list': fitters, 'Welder list': welders}


def return_sorted_and_ranked(df, ei, array_of_ids, col_name, defect_log, state, start_date, end_date, void_entries_with_invalid_employee_number=True):

    
    
    ei_copy = ei.copy()
    # ei_copy['Name'] = ei_copy['<FIRSTNAME>'] + ' ' + ei_copy['<LASTNAME>']
    ei_names = ei_copy[['ID','Name']]
    
    
    # create a dataframe out of the employee IDs
    employees = pd.DataFrame(columns=['ID'], data=array_of_ids)
    employees = employees.set_index('ID').join(ei_names.set_index('ID'), how='inner')
    employees = employees.reset_index(drop=False)
    employees = employees.rename(axis=1, mapper={'index':'ID'})
    
    ''' if you get the error:
        
        ValueError: Length of values does not match length of index
        
        the best way to troubleshoot is to to compare the ids in the 
        'troubleshooter' list to the ids in 'employees' (or 'fitters'/'welders')
        and then find where in the fab listing google sheet that the typo'ed 
        employee number is at. 'troubleshooter' and 'fitter'/'welder' will not 
        be the same length so find the values that are missing
        
        The variables to be assigned are:
            df : this should already be defined in the states-loop
            array_of_ids : let this be either the variable 'fitters' or 'welders'
            col_name : either 'Fitter' or 'Welder'
            defect_log : should already be defined in the states-loop
            state : should already be defined in the states-loop
            start_date : should already be defined in the states-loop
            end_date : should already be defined in the states-loop
        
        '''
    # create an empty list to store names in
    names = []
    troubleshooter = []
    # iterate through the ids in the array_of_ids
    for id_num in array_of_ids:
        # get the row that partains to the employee id number
        x = ei[ei['ID'] == id_num]
        # verify that there is data in that row
        if not x.empty:
            # combine the first and last names
            name = x['Name'].iloc[0]
            # append the name to the list of names
            names.append(name)
            troubleshooter.append(int(id_num))
        else:
            print(id_num)
    
   
    if len(array_of_ids) != len(troubleshooter):
        print(col_name)
        difference = list(np.setdiff1d(array_of_ids, np.array(troubleshooter)))
        print('{} has the following extra IDS: {}'.format(state, difference))
        indexes = list(df[df[col_name].isin(difference)].index)
        print('The index in df with these IDS are: {}'.format(indexes))
        df_to_fix = df.loc[indexes]
        df_to_fix_timestamp = list(df_to_fix['Timestamp'].dt.strftime('%d/%m/%Y %h:%m%s'))
        print('The datetimes for the rows that need fixing are:\n{}'.format(df_to_fix_timestamp))

    ''' This is the old way of doing things and could not be done unless the extra employee ids found in the 'difference'
        variable get fixed 
        Now I just use a join b/c i am a little less of a dummy
        signed -Cody 2022-09-23'''
        
        
    # get the list of names and put it in a column
    # employees['Name'] = names
    
    
    # get the unique jobs in the dataframe
    fab_listing_jobs = pd.unique(df['Job #'])
    # create an empty dict for storing lists of data for each job/employee
    job_weights = {}
    for job in fab_listing_jobs:
        # fill the job_weights dict with keys/lists
        job_weights[str(job)] = []
        
    # initialize empty lists that get appended to for easier assignment to dataframe columns later
    # quantity_list, weight_list, defect_quantity_list, defect_unique_list = [], [], [], []

    
    # set the index as the ID so that the groupby can easily be joined
    employees = employees.set_index('ID')
    # set the classification to the col_name wihch is 'Fitter' or 'Welder'
    employees['Classification'] = col_name
    # set the state location
    employees['Location'] = state
    # groups by the col_name to sum all values and then only keep quantity, weight, earned hours
    df_grouped = df[[col_name, 'Quantity','Weight','Earned Hours']].groupby(col_name).sum()
    # covnert them to numbers?
    df_grouped = df_grouped.apply(pd.to_numeric)
    # join the grouped data onto the  dataframe
    employees = pd.merge(left=df_grouped, right=employees, how='left', left_index=True, right_index=True)
    
    
    # get FIT/WEL to know whether to grab defects that are FIT/WELD, based off the col_name
    employee_type = col_name[:3]      
    if defect_log.shape[0]:
        # only get the columns we are grouping by to avoid errors
        defect_log_to_group = defect_log[['Worked by','Defect Category','Qty.']]
        # only get the pertinent defects categories
        defect_log_to_group = defect_log_to_group[defect_log_to_group['Defect Category'].str.contains(employee_type)]
        # group by employee & category- sum
        defect_log_grouped = defect_log_to_group.groupby(['Worked by']).sum()['Qty.']
        # rename the column to what we wnat it to be called in the employees df
        defect_log_grouped.name = 'Defect Quantity'
        # now join these onto the employees dataframe
        employees = pd.merge(left=employees, right=defect_log_grouped, how='left', left_index=True, right_index=True)
        # set the missing values to zero after our join
        employees['Defect Quantity'] = employees['Defect Quantity'].replace(np.nan, 0)    

        # now we want to ge the unique pieces they defected on?        
        defect_log_grouped = defect_log_to_group.groupby(['Worked by']).count()['Qty.']
        # rename the column to what we wnat it to be called in the employees df
        defect_log_grouped.name = 'Defect Unique'
        # now join these onto the employees dataframe
        employees = pd.merge(left=employees, right=defect_log_grouped, how='left', left_index=True, right_index=True)
        # set the missing values to zero after our join
        employees['Defect Unique'] = employees['Defect Unique'].replace(np.nan, 0)   
        
    else:
        employees['Defect Quantity'] = 0
        employees['Defect Unique'] = 0
 


    
    
    # Convert weight from lbs to tons
    employees['Tonnage'] = employees['Weight'] / 2000    
    # put the quantity of  defects list in as a column
    # calculate average weight per piece
    employees['Weight per Piece'] = employees['Weight'] / employees['Quantity']
    # calculate average tons per piece
    employees['Tonnage per Piece'] = employees['Tonnage'] / employees['Quantity']
    # calculate the number of pieces per defect
    employees['Pieces per Defect'] = employees['Quantity'] / employees['Defect Quantity']
    # calulate the number of tons per defect
    employees['Tons per Defect'] = employees['Tonnage'] / employees['Defect Quantity']
    # These can throw np.inf if someone has no defects
    employees['Pieces per Defect'] = employees['Pieces per Defect'].replace(np.inf, 0)
    # replace any np.inf with zero
    employees['Tons per Defect'] = employees['Tons per Defect'].replace(np.inf, 0)


            
    # group by employee & job number
    df_by_employee_job = df[[col_name, 'Job #', 'Weight']].groupby([col_name,'Job #']).sum()
    # reset index
    df_by_employee_job = df_by_employee_job.reset_index()
    # set the index back to the employee
    df_by_employee_job = df_by_employee_job.set_index(col_name)
    # pivot to get the employee as index and the job as columns and the weight of the job as values in tons
    df_by_employee_job_pivot = df_by_employee_job.pivot(columns='Job #',values='Weight') / 2000
    # repalce nan with zero
    df_by_employee_job_pivot = df_by_employee_job_pivot.fillna(0)
    # rename the columns
    df_by_employee_job_pivot.columns = ['Job:' + str(i) for i in df_by_employee_job_pivot.columns]
    # join onto the employees dataframe
    employees = pd.merge(employees, df_by_employee_job_pivot, how='left', left_index=True, right_index=True)
    
    
    
  
    return employees




def convert_weight_to_earned_hours(state, data_df, drop_job_weights):

    if state == 'TN':
        shop = 'CSM'
    elif state == 'MD':
        shop = 'FED'
    elif state == 'DE':
        shop = 'CSF'



    data_df = data_df.copy()
    
    cols = data_df.columns.tolist()
    weight_cols = []
    # check to see if the column has type 'XXXX Weight' where X=number
    for col in cols:
        try:
            int(col[:4])
            if col[-6:] == 'Weight':
                job = int(col[:4])
                weight_cols.append(col)
        except:
            continue
            
    
    # grab the df columns that are going to be converted to earned hours
    wt_to_earned = data_df[weight_cols]
    # map out the renamed cols
    renamed_cols = {}
    for i in weight_cols:
        renamed_cols[i] = int(i[:4])
    # rename the cols to just the int version of the job number
    wt_to_earned = wt_to_earned.rename(columns=renamed_cols)
    # convert the whole thing to tons
    wt_to_earned = wt_to_earned / 2000
    # get the man hours per ton for that state
    averages = pd.read_excel('c:\\downloads\\averages.xlsx')
    
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
    
    
    
    
    # get the unique jobs from the columns
    jobs = wt_to_earned.columns.tolist()
    earned_from_wt = pd.DataFrame(columns=jobs)
    for job in jobs:
        earned_from_wt[job] = wt_to_earned[job] * averages.loc[job,'Hours per Ton']
    
    earned_from_wt_total = earned_from_wt.sum(axis=1)
    earned_from_wt_total = earned_from_wt_total.rename('Earned Hours')
    
    data_df['Earned Hours 2'] = earned_from_wt_total
    
    if drop_job_weights:
        data_df = data_df.drop(columns=weight_cols)

    return data_df


def combine_multiple_all_both_csv_files_into_one_big_one(list_of_files_to_combine, new_file_output_fullpath_and_name):
    ''' Here below lies the effort to combine multiple csv files to create one giant report '''
    
    
    combo_csv = pd.DataFrame()
    
    for f in list_of_files_to_combine:
        
        this_csv = pd.read_csv(f, index_col=0)
        
        if not combo_csv.shape[0]:
            combo_csv = this_csv
        else:
            combo_csv = combo_csv.append(this_csv, ignore_index=True)
        
    combo_csv_grouped = combo_csv.groupby(by=['Name','Classification','Location']).sum()
    combo_csv_grouped = combo_csv_grouped.reset_index()[['Name','Classification','Location',
                                           'Quantity','Weight','Earned Hours',
                                           'Defect Quantity','Defect Unique', 
                                           'Tonnage','Direct Hours','Indirect Hours',
                                           'Total Hours']]
    combo_csv_grouped = combo_csv_grouped[combo_csv_grouped['Total Hours'] > 0]
    combo_csv_grouped['Weight per Piece'] = combo_csv_grouped['Weight'] / combo_csv_grouped['Quantity']
    combo_csv_grouped['Tonnage per Piece'] = combo_csv_grouped['Tonnage'] / combo_csv_grouped['Quantity']
    combo_csv_grouped['Pieces per Defect'] = combo_csv_grouped['Quantity'] / combo_csv_grouped['Defect Quantity']
    combo_csv_grouped['Pieces per Defect'] = combo_csv_grouped['Pieces per Defect'].replace(np.inf, 0)
    combo_csv_grouped['Tons per Defect'] = combo_csv_grouped['Tonnage'] / combo_csv_grouped['Defect Quantity']
    combo_csv_grouped['Tons per Defect'] = combo_csv_grouped['Tons per Defect'].replace(np.inf, 0)
    combo_csv_grouped['Direct/Total'] = combo_csv_grouped['Direct Hours'] / combo_csv_grouped['Total Hours']
    combo_csv_grouped['DL Efficiency'] = combo_csv_grouped['Earned Hours'] / combo_csv_grouped['Direct Hours']
    combo_csv_grouped['TTL Efficiency'] = combo_csv_grouped['Earned Hours'] / combo_csv_grouped['Total Hours']
    combo_csv_grouped['DL Hrs/Ton'] = combo_csv_grouped['Direct Hours'] / combo_csv_grouped['Tonnage']
    combo_csv_grouped['TTL Hrs/Ton'] = combo_csv_grouped['Total Hours'] / combo_csv_grouped['Tonnage']
    combo_csv_grouped['Earned Hours Rank'] = combo_csv_grouped.groupby(by=['Classification','Location'])['Earned Hours'].rank('first', ascending=False)
    
    combo_csv_grouped.to_csv(new_file_output_fullpath_and_name)
