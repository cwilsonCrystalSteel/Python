# -*- coding: utf-8 -*-
"""
Created on Sun Jul 10 09:10:48 2022

@author: CWilson
"""


import os
from pathlib import Path
import datetime
import pandas as pd
from csv import reader

change_folder = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'Change_Logs'
error_folder = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'Error_logs'



now = datetime.datetime.now()
today = now.date()
start_date = today- datetime.timedelta(days=10)



def daily_errors_summary():
    
    error_files = os.listdir(error_folder)
    todays_changes = [i for i in error_files if datetime.datetime.strptime(i[-23:-10], '%Y-%m-%d-%H').date() >= start_date]
    filename = ''    
    # filename = todays_changes[0]
    errors_df = pd.DataFrame()
    
    
    for filename in todays_changes:
        error_type = filename.split('-')[1][1:-5].split(' for ')[0]
        
        if error_type in ['Invalid Lots Name', '', 'Cleaning Error', 'Retrieval Error']:
            continue
        
        with open(error_folder + '\\' + filename) as file:
            contents = file.readlines()
            contents = [line.rstrip() for line in contents]
            
            pm = filename.split('-')[1][1:-5].split(' for ')[1]

            
            if pm == '':
                pm = 'No PM'
            headers = contents[1].split(',')
            rows = contents[2:]
            rows = [i for i in rows if len(i) > 0]
            
            file_df = pd.DataFrame(columns=headers)
            for line in reader(rows):
                # print(line)
                file_df.loc[len(file_df)] = line
            
            
            
            file_df['filename'] = filename
            file_datetime = datetime.datetime.strptime(filename[-23:-7], "%Y-%m-%d-%H-%M")
            file_df['file datetime'] = file_datetime
            file_df['PM'] = pm
            file_df['Error'] = error_type
            
     
            errors_df = errors_df.append(file_df, ignore_index=True)
                
    
    
    # this is to get rid of anytime an error was pulled before the PM entered their name
    duplicated_work_descs = errors_df[errors_df.duplicated(subset=['Work Description'])]
    duplicated_work_descs = duplicated_work_descs[duplicated_work_descs['PM'] == 'No PM']
    errors_df = errors_df[~errors_df.index.isin(duplicated_work_descs.index)]
    
    # errors_df['file datetime'] = pd.to_datetime(errors_df['file datetime'], errors='coerce') 
    smallest_errors = errors_df.sort_values('file datetime', ascending=False)
    smallest_errors = smallest_errors.drop_duplicates(keep='first', subset=['Job','Fabrication Site','Type of Work','Number'])
    bigger_errors = errors_df.drop_duplicates(keep='first', subset=['Job','Fabrication Site','Type of Work','Number','Work Description','file datetime'])
    bigger_errors = bigger_errors.drop(columns=['Work Description','Delivery','Shipped','PM','Error','filename'])
    
    
    
    
    
    

    
    
    # result = errors_df.groupby(['Job','Fabrication Site','Type of Work','Number']).agg({'file date': ['min', 'max']})
    result = bigger_errors.groupby(['Job','Fabrication Site','Type of Work','Number']).agg(['min', 'max'])
    # this gets rid of the werid double headers caused by the agg function
    result = result.droplevel(0, axis=1)
    # only get errors that are still persistent as of the last round of inspection
    result = result[result['max'] > now - datetime.timedelta(hours=1)]
    if result.shape[0] == 0:
        output2 = pd.DataFrame(columns=['No Errors found today'])
    else:
        
        result['Number of Days Error Present'] = (result['max'] - result['min']).dt.days
        result = result.sort_values('Number of Days Error Present', ascending=False)
        result = result.reset_index(drop=False)
        output = result.merge(smallest_errors, how='left', on=['Job','Fabrication Site','Type of Work','Number'])
        output = output[['PM','Error','Number of Days Error Present','Job','Fabrication Site','Type of Work','Number','Work Description','Delivery','Shipped']]
        
        ''' now I need to go back to the google sheet and get all of the entries for 
        'Duplicate Sequence Number' and put those directly below the one showing up 
        in the output dataframe
        '''
        
        
        count_occurences = bigger_errors.groupby(['Job','Fabrication Site','Type of Work','Number']).count()
        output2 = output.merge(count_occurences, how='left', on=['Job','Fabrication Site','Type of Work','Number'])
        # from the merge, the count column is named 'file datetime'
        output2 = output2.rename(columns={'file datetime':'# of times error found'})
        output2 = output2[['PM','Error','Number of Days Error Present','Fabrication Site','Job','Type of Work','Number','Work Description','Delivery','Shipped','# of times error found']]
    
    
    
    
    return output2



def changes_and_new_work(day_as_date):
    
    
   
    
    change_files = os.listdir(change_folder)
    todays_changes = [i for i in change_files if datetime.datetime.strptime(i[-23:-13], '%Y-%m-%d').date() == day_as_date]
    filename = ''    
    
    changes_df = pd.DataFrame(columns=['today','Shop','Lot','Original Delivery','Previous Delivery','New Delivery','Number Times Changed'])
    new_df = pd.DataFrame(columns=['today','Shop','Type','Name','Delivery Date'])
    
    for filename in todays_changes:
        if os.path.isfile(change_folder + '\\' + filename):
            with open(change_folder + '\\' + filename) as file:
                contents = file.readlines()
                contents = [line.rstrip() for line in contents]
                
                
                filename_type = filename[:-24]
                file_date = filename[-23:-13]
                
                shop = contents[0]
                name = contents[1]
                
                if filename_type == 'Date Change':
                    
                    
                    seqs = contents[4][11:]
                    original_date = contents[5][-10:]
                    changes = contents[6:]
                    date_changes = [i for i in changes if 'Delivery date changed' in i]
                    previous_date = date_changes[-1][-24:-14]
                    new_date = date_changes[-1][-10:]
                    delivery_date_dt = datetime.datetime.strptime(new_date, '%m/%d/%Y').date()
                    if  delivery_date_dt <= day_as_date - datetime.timedelta(days=90):
                        continue                    
                    num_date_changes = len(date_changes)
                    seq_changes = [i for i in changes if 'Sequences changed' in i]
                    num_seq_changes = len(seq_changes)
                    total_num_changes = num_date_changes + num_seq_changes
                    changes_row = {'Shop': shop,
                                   'Lot': name,
                                   'Original Delivery': original_date,
                                   'Previous Delivery': previous_date,
                                   'New Delivery': new_date,
                                   'Number Times Changed': num_date_changes,
                                   'today': file_date}
                    changes_df = changes_df.append(changes_row, ignore_index=True)
                    
                elif 'Added' in filename_type:
                    work_type = contents[2][:contents[2].find(' ')]
                    if work_type not in ['LOT','Ticket','Item','Buyout']:
                        work_type = contents[3][:contents[3].find(' ')]
                        
                    
                    delivery_date = contents[5][-10:]
                    delivery_date_dt = datetime.datetime.strptime(delivery_date, '%m/%d/%Y').date()
                    if  delivery_date_dt <= day_as_date - datetime.timedelta(days=90):
                        continue
                    # sometimes the description is too long and has a line break in it
                    # this causes contents[4] to spill over to contents[5], so the delivery date is then at the end of the next row
                    try:
                        datetime.datetime.strptime(delivery_date, '%m/%d/%Y')
                    except:
                        delivery_date = contents[6][-10:]
                    
                    
                    
                    new_row = {'Shop': shop,
                               'Type': work_type,
                               'Name': name,
                               'Delivery Date': delivery_date,
                               'today': file_date}
                    new_df = new_df.append(new_row, ignore_index=True)
            
    
    
    # filename_types = list(set(filename_types))
    # testing = {key: [] for key in filename_types}
    
    
    if new_df.shape[0]:
        new_df['Delivery Date'] = pd.to_datetime(new_df['Delivery Date'], errors='coerce').dt.date
        new_df = new_df[~new_df['Delivery Date'].isna()]
        new_df['today'] = pd.to_datetime(new_df['today']).dt.date
        new_df['Days until Delivery'] = (new_df['Delivery Date'] - new_df['today']).dt.days
        new_df = new_df.sort_values(by=['Shop', 'Days until Delivery'])
    
    if changes_df.shape[0]:
        changes_df['Original Delivery'] = pd.to_datetime(changes_df['Original Delivery']).dt.date
        changes_df['Previous Delivery'] = pd.to_datetime(changes_df['Previous Delivery']).dt.date
        changes_df['New Delivery'] = pd.to_datetime(changes_df['New Delivery']).dt.date
        changes_df['today'] = pd.to_datetime(changes_df['today']).dt.date
        changes_df['Current Num Days until Delivery'] = (changes_df['New Delivery'] - changes_df['today']).dt.days
        changes_df = changes_df.sort_values(by=['Shop','Current Num Days until Delivery','Lot'])
    
    return {'Changes':changes_df, 'New':new_df}








