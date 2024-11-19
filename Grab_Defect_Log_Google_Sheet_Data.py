# -*- coding: utf-8 -*-
"""
Created on Thu Apr 22 12:30:24 2021

@author: CWilson
"""


import pandas as pd
import gspread
from google_sheets_credentials_startup import init_google_sheet
import datetime 

# state = 'TN'
# start_date = "01/03/2020"
# end_date = "04/30/2021"



def unsplit_shared_pieces_defect_log(df1, col_num, multiple_employee_split_key):
    # get the name of the column based on the column number for simplicity sake
    col_name = df1.columns[col_num]
    # get only pieces with multiple fitters
    shared_pieces = df1[df1[col_name].str.contains(multiple_employee_split_key)]
    
    # don't bother doing anythign if we dont have any shared pieces
    if not shared_pieces.shape[0]:
        print(f'no pieces found with the splitter: {multiple_employee_split_key}')
        return df1
    
    # create a new dataframe which will hold the broken apart entries
    split_pieces = pd.DataFrame(columns=shared_pieces.columns)
    for row in shared_pieces.index:
        # grab that slice of the dataframe
        this_piece = shared_pieces.loc[row].copy()
        # break apart the fitter by the splitting key
        mult_employees = this_piece.loc[col_name].split(multiple_employee_split_key[-1])
        # divide the quantity by the number of people working on it    
        this_piece['Qty.'] = this_piece['Qty.'] / len(mult_employees)
        # iterate through each employee on the piece
        for i,emp in enumerate(mult_employees):
            # this is to get rid of anybody in the Worked by column that is not a number
            try:
                int(emp)
            except:
                print(emp)
                continue
            
            # create a copy of the df/series
            split_piece = this_piece.copy()
            # create the decimal value that will get added to the index to create a unique index
            index_decimal = (i + 1) / 10
            # add the decimal value to the current index (which is the name of the series)
            new_index = split_piece.name + index_decimal
            # round the index to 2 decimal places to eliminate wacky shit like 45.11999999997 instead of 45.12)
            new_index = round(new_index, 1)
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
    df1 = df1[~df1[col_name].str.contains(multiple_employee_split_key)]
    # add in the split_pieces to replace the shared_pieces
    # df1 = df1.append(split_pieces)
    df1 = pd.concat([df1, split_pieces], axis=0, ignore_index=True)
    # puts the pieces back in place where they belong
    df1 = df1.sort_index()
    
    return df1





# def grab_defect_log_OnlyIDS(state, start_date="03/06/1997", end_date="03/06/1997"):
    
#     # get the defect log key for each state, different sheet for different shop
#     if state == 'TN':
#         sheet_key = "1zZLztgYupsglYxGnv140ermvTSolZVe1afr7L-0ctpA"
#     elif state == 'MD':
#         sheet_key = "12EKVlHOe8avAbGf563Uls3JQtRVPndsUUnkruh-sqGM"
#     elif state == 'DE':
#         sheet_key = "1m4v4uCgLbo_pJ2U3BMm4P-L5WBcsX5_BWWDk-bWhXMc"
        
#     # open the Google Sheet
#     sh = init_google_sheet(sheet_key)
#     # The worksheet is always called "Defect Log"
#     worksheet = sh.worksheet("Defect Log")
#     # get all the values as a list of all the rows
#     all_values = worksheet.get_all_values()
#     # convert to a dataframe
#     df = pd.DataFrame(columns=all_values[0], data=all_values[1:])
#     # convert the Date Found column to datetimes
#     df['Date Found'] = pd.to_datetime(df['Date Found'], errors='coerce')
#     # convert the start_date string to a datetime
#     start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
#     # convert the end_date string to a datetime
#     end_dt = datetime.datetime.strptime(end_date, "%m/%d/%Y")
#     # cut out anything before start_date
#     df = df[df['Date Found'] >= start_dt]
#     # cut out anything after end_date
#     df = df[df['Date Found'] <= end_dt]
#     # reset the index, I don't care about the stuff outside of the date range
#     df = df.reset_index(drop=True)
    
    
#     # create empty list to store boolean in
#     has_num_list = []
#     # iterate thru the Worked By column
#     for worker in df['Worked by'].tolist():
#         # checks if any of the values in the string are a number
#         booooolean = any(i.isdigit() for i in worker)
#         # appends the boolean to the list
#         has_num_list.append(booooolean)
#     # drop anything in the df if the Worked By does not contain a number (number = Employee ID)
#     df = df[has_num_list]
#     # convert the QTY column to int
#     df['Qty.'] = df['Qty.'].apply(pd.to_numeric, errors='coerce')
#     # Split up any entries with multiple employees based on the period
#     df = unsplit_shared_pieces_defect_log(df, 6, '\.')
#     # Split up any entries with multiple employees based on the slash
#     df = unsplit_shared_pieces_defect_log(df, 6, '/')
#     # conver the worked by column to ints
#     df['Worked by'] = df['Worked by'].apply(pd.to_numeric)
    
    
#     return df


def grab_defect_log(state, start_date="03/06/1997", end_date="03/06/1997", worked_by_employeeIDs_only=True):
    
    # get the defect log key for each state, different sheet for different shop
    if state == 'TN':
        sheet_key = "1zZLztgYupsglYxGnv140ermvTSolZVe1afr7L-0ctpA"
    elif state == 'MD':
        sheet_key = "12EKVlHOe8avAbGf563Uls3JQtRVPndsUUnkruh-sqGM"
    elif state == 'DE':
        sheet_key = "1m4v4uCgLbo_pJ2U3BMm4P-L5WBcsX5_BWWDk-bWhXMc"
        
    # open the Google Sheet
    sh = init_google_sheet(sheet_key)
    # The worksheet is always called "Defect Log"
    worksheet = sh.worksheet("Defect Log")
    # get all the values as a list of all the rows
    all_values = worksheet.get_all_values()
    # convert to a dataframe
    df = pd.DataFrame(columns=all_values[0], data=all_values[1:])
    # convert the Date Found column to datetimes
    df['Date Found'] = pd.to_datetime(df['Date Found'], errors='coerce')
    # convert the start_date string to a datetime
    start_dt = datetime.datetime.strptime(start_date, "%m/%d/%Y")
    # convert the end_date string to a datetime
    end_dt = datetime.datetime.strptime(end_date, "%m/%d/%Y")
    # cut out anything before start_date
    df = df[df['Date Found'] >= start_dt]
    # cut out anything after end_date
    df = df[df['Date Found'] <= end_dt]
    # reset the index, I don't care about the stuff outside of the date range
    df = df.reset_index(drop=True)
    # convert the QTY column to int
    df['Qty.'] = df['Qty.'].apply(pd.to_numeric, errors='coerce')    
    # create empty list to store boolean in
    has_num_list = []
    # iterate thru the Worked By column
    for worker in df['Worked by'].tolist():
        # checks if any of the values in the string are a number
        booooolean = any(i.isdigit() for i in worker)
        # appends the boolean to the list
        has_num_list.append(booooolean)
    
    # if you are keeping the other 'Worked by' categories then it will grab those from the df
    if not worked_by_employeeIDs_only:
        # get other defects that are not worked on by an employee ID
        df1 = df[[not boool for boool in has_num_list]]      
        
    if len(has_num_list):
        # drop anything in the df if the Worked By does not contain a number (number = Employee ID)
        df = df[has_num_list]
    # Split up any entries with multiple employees based on the period
    df = unsplit_shared_pieces_defect_log(df, 6, '\.')
    # Split up any entries with multiple employees based on the slash
    df = unsplit_shared_pieces_defect_log(df, 6, '/')
    
    df = unsplit_shared_pieces_defect_log(df, 6, '-')
    
    df = unsplit_shared_pieces_defect_log(df, 6, ';')
    
    df = unsplit_shared_pieces_defect_log(df, 6, ' ')

    # convert employee IDs to numbers
    df['Worked by'] = df['Worked by'].apply(pd.to_numeric)
    # if you are keeping the other 'Worked by' categories then it will append those to the main df
    if not worked_by_employeeIDs_only:
        # add the pieces that are not worked by employee IDs back to the df
        # df = df.append(df1)
        df = pd.concat([df, df1], axis=0)
        
    # sort it by the index
    df = df.sort_index(ascending=True)
    
    return df














