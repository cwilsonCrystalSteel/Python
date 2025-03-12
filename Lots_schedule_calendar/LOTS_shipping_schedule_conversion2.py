# -*- coding: utf-8 -*-
"""
Created on Thu Jun  3 08:18:31 2021

@author: CWilson
"""
#%% Functions & Importing Modules
import pandas as pd
import datetime
import glob
import os
from pathlib import Path
from utils.attendance_google_sheets_credentials_startup import init_google_sheet
from Lots_schedule_calendar.calendar_error_producer_function import produce_error_file
from Lots_schedule_calendar.calendar_emailing_functions_with_gmail_api import send_error_notice_email

pm_emails = {'Darryl': 'dgassaway@crystalsteel.net',
             'Darryl ': 'dgassaway@crystalsteel.net',
             'Mike D': 'mdorsey@crystalsteel.net',
             'Jon T': 'cwilson@crystalsteel.net',
             'Joel': 'jnewsom@crystalsteel.net',
             'John S': 'jshubbuck@crystalsteel.net',
             'Anthony B': 'cwilson@crystalsteel.net',
             'Mustafa': 'mkilicarslan@crystalsteel.net',
             'Patrick K': 'pklein@crystalsteel.net',
             'Ahmed': 'cwilson@crystalsteel.net',
             'Scott': 'sundershute@crystalsteel.net',
             'Joel ': 'jnewsom@crystalsteel.net',
             'AN/JN': 'cwilson@crystalsteel.net',
             'AN': 'cwilson@crystalsteel.net',
             'Ahmed/Joel': 'jnewsom@crystalsteel.net',
             'Dick G': 'rgast@crystalsteel.net',
             '': 'cwilson@crystalsteel.net',
             'Ahmed/Mustafa': 'mkilicarslan@crystalsteel.net',
             'Mustafa ': 'mkilicarslan@crystalsteel.net',
             'Jessica\nJohn S.': 'jshubbuck@crystalsteel.net',
             'John S.': 'jshubbuck@crystalsteel.net',
             'Mark K (CMW)': 'cwilson@crystalsteel.net',
             'Joel\nJoel': 'jnewsom@crystalsteel.net',
             'Patrick': 'pklein@crystalsteel.net',
             'Mishler': 'cwilson@crystalsteel.net'}


manila_lots_emails = ['nmarinduque@crystalsteel.net','rrivera@crystalsteel.net', 
                      'aanonuevo@crystalsteel.com']

_GoogleSheetMagicKeyThatIDontRemeberWhatItIsActuallyCalled = "1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4"


def get_shipping_schedule(shop, type_of_work, send_emails=False):
    try:
        # get the production worksheet 
        sh = init_google_sheet(_GoogleSheetMagicKeyThatIDontRemeberWhatItIsActuallyCalled)
        # get the values from the shipping schedule as a list of lists
        worksheet = sh.worksheet('Shipping Sched.').get_all_values()
        
        cols_to_keep = ['Job',
                         'Fabrication Site',
                         'Type of Work',
                         'Number',
                         'Tonnage',
                         'Estimated Hours for Ticket',
                         'Delivery',
                         'Work Description',
                         'Complete?',
                         'Shipped?',
                         'Dispatcher Notes',
                         'PM']
        i = 0
        while i < 20:
            i += 1
            column_failure = True
        
            # convert to a dataframe. row 2 is columns, dont care about stuff before row 10
            ss = pd.DataFrame(worksheet[i+1:], columns=worksheet[i])
            # empty dict to get rid of line breaks in the column names
            new_cols = {}
            for col in ss.columns:
                new_col = col.replace('\n', ' ')
                new_cols[col] = new_col
                
            # replace columns with new columns w/o line breaks
            ss = ss.rename(columns=new_cols)
            ss['GoogleSheetRowNumber'] = [i + 2 + x for x in list(range(0,ss.shape[0]))]
            # only keep things marked as sequence
            try:
                print('\n' + str(i), end='\t')
                ss = ss[ss['Type of Work'] == type_of_work]
                
                # count how many of our columns are present in the dataframe colummns
                num_cols_in_df = len([i for i in cols_to_keep if i in ss.columns])
                # if there are more than half that match, then proceed to the infill of missing columns
                if num_cols_in_df > (len(cols_to_keep)+1)/2:
                    # this infilss missing columns based on indexes - which is a bad bandaid fix
                    for xyz, col in enumerate(cols_to_keep):
                        # print(xyz, col)
                        if col not in ss.columns:
                            ss.columns.values[xyz] = col
                        
                # only keep the columns we want 
                ss = ss[['Job','Fabrication Site', 'Type of Work','Number','Work Description','Delivery', 'PM', 'Shipped?', 'Dwgs Needed in Shop', 'GoogleSheetRowNumber']]
            
                column_failure = False
                print('yay')
                break
            except:
                print('whaaa')
                continue
            
        # this is the backup if the while loop can't find the header row
        if column_failure:
            # define what the columns *SHOULD* be
            backup_cols = ['Job', 'Fabrication Site', 'Type of Work', 
                            'Number', 'Delivery','Work Description',
                            'Shipped?','Dispatcher Notes','PM','Tonnage','Drawing Release %',
                            'Purchasing Start-Buying Sched','Percent of material received',
                            'Fab % Completed by Sequence','Number of Main Members',
                            "Detailing PM's Comments", '% of Dwgs W/O Pending Issues','Drawings Submitted for approval by','Actual approval date',
                            'Dwgs Needed in Shop', 'Actual issue date of fabrication drawings', 'Ship to paint or Galv',
                            'Detailing', 'Joist and Deck Detailing and Supply',
                            'In house or subcontract engineering', 'Supply of Raw materails',
                            'Consumables and shop supplies', 'Erection',
                            'Specialty Buyouts', 'Notes', 'Fab tonnage left',
                            '', '', '', '', '', '','','']
            # chop the df so it is only the same number of columns as the provided number of columns
            ss = ss.iloc[:,0:len(backup_cols)]
            # rename the columns 
            ss.columns = backup_cols
            # keep only the shit we want
            ss = ss[['Job','Fabrication Site', 'Type of Work','Number','Work Description','Delivery', 'PM', 'Shipped?']]
            # get the google sheet row number
            ss['GoogleSheetRowNumber'] = [i + 2 + x for x in list(range(0,ss.shape[0]))]
            # get the type of work we are after
            ss = ss[ss['Type of Work'] == type_of_work]
            
            # chop off the first 100 rows - likely won't need those 
            ss = ss.iloc[100:,:]
        else:
            print('found the column headers on row ' + str(i))
                
                
                
        if shop != None:
            # only get csm 
            ss = ss[ss['Fabrication Site'] == shop]
            ss['PM'] = ss['PM'].str.replace('\n', '--')
        
        try:
            # ss['Shipped?'][ss['Shipped?'] != "Yes"] = ''
            ss['Shipped?'] = ss['Shipped?'].replace(['No','Partial'], '')
        except:
            print('Could not fix the shipped column')
       
        return ss
    
    except Exception as e:
        error_name = 'Shipping Schedule - Retrieval Error'
        produce_error_file(exception_as_e = e, shop = shop, file_prefix = error_name)
        if send_emails:
            send_error_notice_email(shop, error_name, e)        

    

def apply_google_sheets_link(ss, column_name=None):
    column_name_to_cell_column_letter_dict = {'Job': 'A',
                                              'Fabrication Site': 'B',
                                              'Type of Work': 'C',
                                              'Number': 'D',
                                              'Delivery': 'G',
                                              'Work Description': 'H',
                                              'Shipped?': 'J',
                                              'PM': 'L',
                                              'Dwgs Needed in Shop': 'T'}
    
    if column_name:
        CELL_COLUMN_LETTER = column_name_to_cell_column_letter_dict[column_name]
    else:
        CELL_COLUMN_LETTER = 'D'
    # 'https://docs.google.com/spreadsheets/d/1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4/edit#gid=396215481&range=A235'

    url_start = 'https://docs.google.com/spreadsheets/d/' + _GoogleSheetMagicKeyThatIDontRemeberWhatItIsActuallyCalled
    url_end = "/edit#gid=396215481&range=" + CELL_COLUMN_LETTER
    
    ss['GoogleSheetLink'] = url_start + url_end + ss['GoogleSheetRowNumber'].astype(str)
    
    
    return ss
    
    
    
    

def shipping_schedule_cleaner_of_duplicates(ss, shop, send_emails=False):
    ''' This finds out if there are duplicate sequence numbers in the 'Number' 
        column. It then removes those values and returns the shipping 
        schedule. After that it generates an error log file & sends an email out
    '''
    # thesse are the duplicates when multiple sequences are present
    dupes = ss[ss.duplicated(subset=['Job','Number'], keep=False)]
    # order the duplicates for the email
    dupes = dupes.sort_values(by=['Job','Number'])
    # get rid of any rows that have alrady shipped
    dupes = dupes[dupes['Shipped?'] == '']
    # THIS GETS THE duplicates IN CASE THE DUPLICATE OF A NOt SHIPPED HAS SHIPPED
    dupes_plus_shipped = ss[ss['Job'].isin(dupes['Job']) & ss['Number'].isin(dupes['Number'])]
    # sort for viewing pleasure 
    dupes_plus_shipped = dupes_plus_shipped.sort_values(by='Number')
    
    # remove those duplicates
    ss = ss[~ss.index.isin(dupes.index)]
    
    # send email for each PM in the list of errors
    for pm in pd.unique(dupes_plus_shipped['PM']):
        dupes_pm = dupes_plus_shipped[dupes_plus_shipped['PM'] == pm]
        # only generate error file & send email if there are entries in dupes_plus_shipped
        if dupes_plus_shipped.shape[0] != 0:
            try:
                pm_email = pm_emails[pm]
            except:
                pm_email = 'cwilson@crystalsteel.net'
            error_name = 'Shipping Schedule - Duplicate Sequence Numbers for ' + pm
            produce_error_file(dupes_pm.to_csv(index=False), shop, file_prefix=error_name)
            if send_emails:
                send_error_notice_email(shop, error_name, dupes_pm.to_html(index=False), extra_recipient=pm_email)
        
    
    return ss
    


def shipping_schedule_cleaner_of_bad_number_column(ss, shop, send_emails=False):
    ''' This finds out if there are values in the 'Number' column that are not 
        actually numbers. It then removes those values and returns the shipping 
        schedule. After that it generates an error log file & sends an email out
    '''
    # copy the shipping schedule
    ss2 = ss.copy()
    # try and convert the number col to float/number
    ss2['Number'] = pd.to_numeric(ss2['Number'].copy(), errors='coerce')
    # get the rows that would not convert
    nans = ss2[ss2['Number'].isna()]
    # get the original rows from the original shipping schedule to get the original number value
    nans = ss.loc[nans.index]
    # get rid of the records that have already shipped
    nans = nans[nans['Shipped?'] == '']
    # remove those rows from the df & return the cleaned shipping schedule
    ss = ss[~ss.index.isin(nans.index)]
    
    # ss = ss.copy()
    
    # send email for each PM in the list of errors
    for pm in pd.unique(nans['PM']):
        nans_pm = nans[nans['PM'] == pm]
        # only generate error file & send email if there are entries in nans
        if nans.shape[0] != 0:
            try:
                pm_email = pm_emails[pm]
            except:
                pm_email = 'cwilson@crystalsteel.net'            
            error_name = 'Shipping Schedule - Sequence Number Not A Number for ' + pm
            produce_error_file(nans_pm.to_csv(index=False), shop, file_prefix=error_name)
            if send_emails:
                send_error_notice_email(shop, error_name, nans_pm.to_html(index=False), extra_recipient=pm_email)    
    
    return ss

def shipping_schedule_cleaner_of_bad_dates(ss, shop, send_emails=False):
    ''' I am doing this because I want the dates in ss2['Delivery'] to be checked each row 
    I cannot trust these people to enter consistent date formatting'''
    import warnings
    warnings.simplefilter(action='ignore', category=UserWarning)
    
    ss2 = ss.copy()
    # convert delivery to a datetime
    ss2['Delivery'] = pd.to_datetime(ss2['Delivery'].copy(), errors='coerce').dt.date
    # get those rows with dates that dont convert from ss2
    bad_dates = ss2[ss2['Delivery'].isna()]
    # get the originial rows of the bad dates
    bad_dates = ss.loc[bad_dates.index]
    # get rid of the rows that have already shipped
    bad_dates = bad_dates[bad_dates['Shipped?'] == '']
    # drop the rows with bad dates
    ss = ss[~ss.index.isin(bad_dates.index)]
    
    ss = ss.copy()
    # convert the remaining items to a date
    ss['Delivery'] = pd.to_datetime(ss['Delivery'].copy(), errors='coerce').dt.date
    
    for pm in pd.unique(bad_dates['PM']):
        bad_dates_pm = bad_dates[bad_dates['PM'] == pm]
        # only generate errors if there are bad dates present
        if bad_dates.shape[0] != 0:
            try:
                pm_email = pm_emails[pm]
            except:
                pm_email = 'cwilson@crystalsteel.net'            
            error_name = 'Shipping Schedule - Delivery Date Not A Date for ' + pm
            produce_error_file(bad_dates_pm.to_csv(index=False), shop, file_prefix=error_name)
            if send_emails:
                send_error_notice_email(shop, error_name, bad_dates_pm.to_html(index=False), extra_recipient=pm_email)    
    
    return ss        



def clean_shipping_schedule(ss, shop, send_emails=False):
    try:
        # get rid of the easy job-seq combos in the lots log
        ss = ss.copy()
        # convert job to a number - if not a number then make it nan
        ss['Job'] = pd.to_numeric(ss['Job'].copy(), errors='coerce')
        # get rid of those shitty inputs in the job number col
        ss = ss[~ss['Job'].isna()]        
        # ss[ss['Job'] == 2038][ss['Number'] == '10']
        # get rid of the items that are duplicates (ex. Sequence 24 is repeated)
        ss = shipping_schedule_cleaner_of_duplicates(ss, shop, send_emails)
        # get rid of the items that have a bad number column & inform me of it
        ss = shipping_schedule_cleaner_of_bad_number_column(ss, shop, send_emails)
        # get rid of the items with bad dates
        ss = shipping_schedule_cleaner_of_bad_dates(ss, shop, send_emails)

        return ss
    
    except Exception as e:
        error_name = 'Shipping Schedule - Cleaning Error'
        produce_error_file(exception_as_e = e, shop = shop, file_prefix = error_name) 
        if send_emails:
            send_error_notice_email(shop, error_name, e)




def get_open_lots_log(shop, send_emails=False):
    try:
        sh = init_google_sheet("1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4")
        worksheet = sh.worksheet('LOTS Log').get_all_values()
        # covnert it to a dataframe
        ll = pd.DataFrame(worksheet[3:], columns=worksheet[2])
        # empty dict to get rid of line breaks in the column names
        new_cols = {}
        for col in ll.columns:
            new_col = col.replace('\n', ' ')
            new_cols[col] = new_col
        # rename the columns with the new columns without line breaks
        ll = ll.rename(columns=new_cols)
        # only keep the columns i want
        ll = ll[['Job','Seq. #','Fabrication Site', 'LOTS Name', 'Fab Release Date', 'Transmittal #', 'Delivery Date']]
        # rename the delivery date column
        ll = ll.rename(columns={'Delivery Date':'LL Delivery Date'})
        # only get the lots that have a fab release date 
        ll = ll[~(ll['Fab Release Date'] == '')]
        # only get this shops lots
        ll = ll[ll['Fabrication Site'] == shop]
        
        return ll
    
    except Exception as e:
        error_name = 'LOTS Log Retrieval Error'
        produce_error_file(exception_as_e = e, shop = shop, file_prefix = error_name)
        if send_emails:
            send_error_notice_email(shop, error_name, e)        



def lots_log_cleaner_of_bad_job_number(ll, shop, send_emails=False):
    ll2 = ll.copy()
    # convert the job column to numbers
    ll2['Job'] = pd.to_numeric(ll2['Job'].copy(), errors='coerce')
    # get the records where the job # wont convert
    bad_jobs = ll2[ll2['Job'].isna()]
    # get the values of the job numbers column from the original df
    bad_jobs = ll[ll.index.isin(bad_jobs.index)]
    # ??? wtf do i go back to using ll instead of ll2 & dropping out the indexes from bad_jobs???
    ll['Job'] = pd.to_numeric(ll['Job'].copy(), errors='coerce')
    # create errors if there are records in the bad_jobs df
    if bad_jobs.shape[0] != 0:
        error_name = 'Lots Log - Lot Job Number Not A Number'
        produce_error_file(bad_jobs.to_csv(index=False), shop, file_prefix=error_name)
        if send_emails:
            send_error_notice_email(shop, error_name, bad_jobs.to_html(index=False), manila_lots_emails)   
    
    return ll
  




  
def lots_log_cleaner_of_bad_dates(ll, shop, send_emails=False):
    ll2 = ll.copy()
    # convert delivery to a datetime
    ll2.loc[:,'Fab Release Date'] = pd.to_datetime(ll2['Fab Release Date'].copy(), errors='coerce').dt.date
    ll2.loc[:,'LL Delivery Date'] = pd.to_datetime(ll2['LL Delivery Date'].copy(), errors='coerce').dt.date
    # get those rows with dates that dont convert from ll2
    bad_dates = ll2[ll2['Fab Release Date'].isna()]
    # get the originial rows of the bad dates
    bad_dates = ll.loc[bad_dates.index]
    # drop the rows with bad dates
    ll = ll[~ll.index.isin(bad_dates.index)]
    # convert the remaining items to a date
    ll.loc[:,'Fab Release Date'] = pd.to_datetime(ll['Fab Release Date'].copy(), errors='coerce').dt.date
    ll.loc[:,'LL Delivery Date'] = pd.to_datetime(ll['LL Delivery Date'].copy(), errors='coerce').dt.date
    # only generate errors if there are bad dates present
    if bad_dates.shape[0] != 0:
        error_name = 'Lots Log - Fab Release Date Not A Date'
        produce_error_file(bad_dates.to_csv(index=False), shop, file_prefix=error_name)
        if send_emails:
            send_error_notice_email(shop, error_name, bad_dates.to_html(index=False), manila_lots_emails)  
    return ll       
    


def lots_log_cleaner_of_invalid_lots_name(ll, shop, send_emails=False):
    # check if the Lot's name is long enough
    ll2 = ll.copy()
    ll2 = ll2[ll2['LOTS Name'].str.len() < 9]
    # check if the Lot number is just zeros
    ll3 = ll.copy()
    ll3 = ll3[ll3['LOTS Name'].str[6:9] == '000']
    # check if the first 4 of the lot name is the job number or not
    ll4 = ll.copy()
    ll4 = ll4[~ll4['LOTS Name'].str[:4].str.isnumeric()]    
    # add up all of ll2, ll3, and ll4
    bad_names = pd.concat([ll2, ll3, ll4], axis=0)
    # get rid of those lots with invalid names
    ll = ll[~ll.index.isin(bad_names.index)]
    # ONLY generate error if there are invalid lot names still present
    if bad_names.shape[0] != 0:
        error_name = 'Lots Log - Invalid Lots Name'
        produce_error_file(bad_names.to_csv(index=False), shop, file_prefix=error_name)
        if send_emails:
            send_error_notice_email(shop, error_name, bad_names.to_html(index=False), manila_lots_emails)
    
    return ll


def lots_log_convert_sequences_col_to_list(ll):
    ll2 = ll.copy()
    # the pipe symbol (|) represents comma OR ampersand
    ll2['Seq. #'] = ll2['Seq. #'].str.split(',|&')
    
    funky_lots = []
    # go thru each row the dataframe
    for i in ll2.index:
        # get that list of sequences for that row
        row = ll2.loc[i]
        # this will be altered inplace - very confusing & idk how/why i did it this way but it seems to work
        seq_list = row['Seq. #']
        # create a copy 
        # seq_list_old = seq_list.copy()
        # # this is just to show the differences if changes were made
        # toggle=False
        
        # go thru each item in that list
        for j, seq_value in enumerate(seq_list):
            # if it starts or ends with whitespace - strip
            if len(seq_value) and (seq_value[0] == ' ' or seq_value[-1] == ' '):
                # change the value in the list to the new value w/o whitespaces
                seq_list[j] = seq_value.strip()
                # add the lot to the funky_lots list
                funky_lots.append(ll2.loc[i,'LOTS Name'])
                # enable showing the difference in sequence lists b/c a change was made
                toggle=True
            # get rid of any newlines
            if '\n' in seq_value:
                # print('\tremoving newline')
                seq_list[j] = seq_value.replace('\n','')
                # add the lot to the funky_lots list
                funky_lots.append(ll2.loc[i,'LOTS Name'])
                # enable showing the difference in sequence lists b/c a change was made
                toggle=True
        
        # if toggle:
        #     print(row['LOTS Name'])
        #     print('old:\n\t ',seq_list_old)
        #     print('new:\n\t ',seq_list)
    
    funky_lots = list(set(funky_lots))
    print('These lots had their sequences fixed up:')
    print(funky_lots)
    return ll2



def lots_log_cleaner_of_duplicate_lots(ll):
    dupes = ll[ll.duplicated(subset=['LOTS Name'], keep=False)]
    
    condensed = pd.DataFrame()
    for lot in pd.unique(dupes['LOTS Name']):
        chunk = dupes[dupes['LOTS Name'] == lot]
        sequences = chunk['Seq. #']
        
        sequences_list = []
        # each seq should be a list at this point
        for seq in sequences:
            for i in seq:
                sequences_list.append(i)
        
        # only get the first appearance of the duplicated lot name
        output_chunk = chunk.iloc[[0]].copy()
        # set the sequence value ot be the new list of all the sequences
        output_chunk.at[output_chunk.index[0],'Seq. #'] = list(set(sequences_list))
        # append it to the condensed df -> this gets put back onto ll
        condensed = pd.concat([condensed, output_chunk], axis=0)
        
    # get rid of all the duplicated lots names rows 
    ll = ll[~ll.index.isin(dupes.index)]
    # put back the condensed versions 
    ll = pd.concat([ll, condensed], axis=0)
    # put it back in about the same order it started 
    ll = ll.sort_index()
    return ll
    
    
    

def clean_lots_log(ll, shop, send_emails=False):
    try:
        # get rid of anything where the Job # is not valid
        ll = lots_log_cleaner_of_bad_job_number(ll, shop, send_emails)
        # get rid of things without a valid fab release date (b/c it gets used in the calendar)
        ll = lots_log_cleaner_of_bad_dates(ll, shop, send_emails)
        # get rid of invalidly named lots
        ll = lots_log_cleaner_of_invalid_lots_name(ll, shop, send_emails)
        # break the 'seq #' column from a string of sequences to a list of sequences
        # the delimiters being used are a comma (',') and ampersand ('&')
        ll = lots_log_convert_sequences_col_to_list(ll)
        # condense the old-method where a LOT spans multiple rows
        ll = lots_log_cleaner_of_duplicate_lots(ll)
        
        return ll
    
    except Exception as e:
        error_name = 'LOTS Log Cleaning Error'
        produce_error_file(exception_as_e = e, shop = shop, file_prefix = error_name)  
        if send_emails:
            send_error_notice_email(shop, error_name, e, manila_lots_emails)


#%%

# ss_open = get_open_shipping_schedule(shop) # ss = shipping schedule
# ll = get_open_lots_log(shop) # ll = lots log

#types_of_work = ['Sequence','Ticket','Item','Buyout']

def retrieve_from_prod_schedule(shop, type_of_work, send_emails):
    # get all of the items from the shipping schedule for this shop
    ss_all = get_shipping_schedule(shop, type_of_work, send_emails)
    # apply the google sheets link
    ss_all = apply_google_sheets_link(ss_all)
    
    # clean it up -> remove duplicate sequences, bad numbers, bad dates
    ss_all = clean_shipping_schedule(ss_all, shop, send_emails)
    
    return ss_all

    

def draw_the_rest_of_the_horse(shop, send_emails=False):
    # get all of the items from the shipping schedule for this shop
    ss_all = retrieve_from_prod_schedule(shop, 'Sequence', send_emails)
    # get the shipping schedule for the shop
    ss_open = ss_all[ss_all['Shipped?'] == '']
    
    
    # clean the shipping schedule
    # ss_open = clean_shipping_schedule(ss_open, shop)
    # get the lots log for the shop
    ll = get_open_lots_log(shop, send_emails)
    # clean the lots log shop
    ll = clean_lots_log(ll, shop, send_emails)

    
    if ll is None:
        print(shop)
    
    open_lots = {}
    
    lots = pd.unique(ll['LOTS Name'])
    
    for lot in lots:
        # print(lot)
        chunk = ll[ll['LOTS Name'] == lot].squeeze()
        # pulling values out as variables to make life easy
        job = chunk['Job']
        sequences = chunk['Seq. #']
        if isinstance(sequences, str):
            sequences = [sequences]
        fab_release_date = chunk['Fab Release Date']      
        ll_delivery_date = chunk['LL Delivery Date']
        
        # First get only that jobs data
        ss_open_job = ss_open[ss_open['Job'] == job]
        
        
        
        # then limit it to only the sequences present in the lot
        ss_open_job = ss_open_job[ss_open_job['Number'].isin(sequences)]

        # skip the rest of the loop if there is nothing left in ss_open_job
        if ss_open_job.shape[0] == 0:
            # print('No open sequences for: ' + lot)
            continue
       
        # get all of the sequence information from the complete shipping schedule
        ss_job_all = ss_all[(ss_all['Job'] == job) & (ss_all['Number'].isin(sequences))]
        # get the earliest delivery date
        if isinstance(ll_delivery_date, datetime.date) and not pd.isnull(ll_delivery_date):
            chosen_delivery = ll_delivery_date
            
            if ll_delivery_date != min(ss_job_all['Delivery']):
                print('LOTS Log date and the earliest Sequence Date do not match: ' + lot)
        else:
            chosen_delivery = min(ss_job_all['Delivery'])
            
        
        try:
            ss_open_job_earliest = ss_open_job[ss_open_job['Delivery'] == chosen_delivery]
            # if there is no entry in the shipping schedule with that delivery date, then throw error which causes to get just the earliest delivery date row
            if not ss_open_job_earliest.shape[0]:
                THROWING_AN_ERROR_do_i_needs_to_get_hyperlink_from_LOTS_log_questionMark
        except Exception:
            ss_open_job_earliest = ss_open_job.sort_values('Delivery')
            
        url = ss_open_job_earliest['GoogleSheetLink'].iloc[0]

        comment = ''
        delta = (chosen_delivery - fab_release_date).days
        if delta <= 0:
            comment += 'Delivery scheduled prior to Release\n'
        elif delta <=7:
            comment += 'Delivery within 7 days of Release\n'
        
        
        
        # add things to the dict
        open_lots[lot] = {}
        open_lots[lot]['Job'] = job
        open_lots[lot]['Sequences'] = ", ".join(sequences)
        open_lots[lot]['Earliest Delivery'] = chosen_delivery
        open_lots[lot]['Fab Release Date'] = fab_release_date
        open_lots[lot]['Comment'] = comment
        open_lots[lot]['URL'] = url

        
        

    if len(open_lots):
        open_lots_df = pd.DataFrame().from_dict(open_lots, orient='index')
        open_lots_df = open_lots_df.reset_index(drop=False)
        open_lots_df = open_lots_df.sort_values(by=['Earliest Delivery','Job'])
        open_lots_df = open_lots_df.set_index('index')
        open_lots_df = open_lots_df.drop(columns=['Job'])
    
    
        today = datetime.datetime.today()
        # todays date string for the excel file output
        today_str = today.strftime("%m-%d-%Y")
        folder_path = Path.home() / 'documents' / 'LOT_schedule_dump'
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        # send the lots_df to excel file
        outfile = folder_path / ('todays_lot_info_' + today_str + ' ' + shop + '.xlsx')
        open_lots_df.to_excel(outfile)
    
    