# -*- coding: utf-8 -*-
"""
Created on Thu Apr  3 14:59:28 2025

@author: Netadmin
"""
''' get_model_estimate_hours_attached_to_fablisting_SQL '''
from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
import pandas as pd
from utils.initSQLConnectionEngine import yield_SQL_engine
from utils.sql_print_count_results import print_count_results, table_exists
from pathlib import Path
import re

engine = yield_SQL_engine()


def further_fablisting_df_preperations(fablisting_df): # move this to get_model_estimate_hours_attached_to_fablisting_SQL
    df = fablisting_df.copy()
    
    df['pcmark'] = df['Piece Mark - REV'].apply(lambda x: x.split('-')[0])
    df['rev'] = df['Piece Mark - REV'].apply(lambda x: x.split('-')[1] if '-' in x else 0)
    
    df['lot_3_digit'] = df['Lot #'].str.zfill(3)
    df['lot_with_t_start'] = 'T' + df['lot_3_digit']
    
    df['shop'] = df['sheetname'].apply(lambda x: x.split(' ')[0])
    
    cols = df.columns
    new_cols = []
    counter = 0
    for i in cols:
        if i == '' or i.isspace() or not i:
            new_cols.append(f'column {counter}')
            counter += 1
        else:
            new_cols.append(i)
    
    df.columns = [i.lower() for i in new_cols]
            
    # df['Job #'] = df['Job #'].astype(str)
    
    return df

def lotnumber_cleaner(lot): # move this to get_model_estimate_hours_attached_to_fablisting_SQL
    if 'ph' in lot.lower():
        out = 'use job average' # cant do anything with these
        return out 



    isticket = False
    islot = False
    
    # I pulled all of fablisting for all shops and did :
    # x = list(set(fablisting['Lot #']))
    # x = [i for i in x if not 'ph' in i.lower()]
    # x = [re.split(r'(\d+)', re.sub(r'[ -#-]', '', i.upper())) for i in x]
    # prefix_options = list(set([i[0] for i in x]))
    # then got all the ones that look like ticket or its abbreviations: 
    existing_ticket_options = ['TK', 'TKT', 'TCK', 'TCT', 'WTK', 'WKTK', 'TICKET', 'WKT', 'TKK']
    # make a regex pattern out of those values 
    ticket_pattern =  '(' + '|'.join(existing_ticket_options) + ')'
    
    # remove anything not a letter or number
    # cleaned = re.sub(r'[^A-Z0-9.]', '', lot.upper())
    
    # First, replace non-alphanumeric separators (like '-') with a space
    lot_spaced = re.sub(r'[^A-Z0-9]', ' ', lot.upper())
    
    # Now remove spaces but preserve number separation
    cleaned_parts = lot_spaced.split()  # Splits into separate parts
    cleaned = ''.join(cleaned_parts)  # Reconstruct without merging numbers

    
    # if we clean it and there is nothing left!
    if not len(cleaned):
        print('No substance to the lot!')
        return None
    
    # if we clean it and then there are no numbers!
    if not re.search(r'[0-9]', cleaned):
        print('No numbers to the lot!')
        return None
    
    # check if its a ticket 
    if re.match(ticket_pattern, cleaned_parts[0]):
        print(f'{lot} is a ticket')
        isticket = True
    
    # if we didnt get a match on the ticket options
    else:
        # first lets just try and make it a number
        try:
            number = str(int(float(cleaned)))
            # if that succeeded, lets move on
            islot = True
        
        # if we had things other than numbers, lets try and pull out any letters
        except ValueError:
            try:
                # try splitting on alphabet values
                alpha_split = re.split(r'[A-Z]', cleaned)
                # if we split on alphabets on the start/end, we will get blank strings, so remove those, and spaces
                alpha_split = [i for i in alpha_split if not re.match(r'(\s)|(^$)',i)]
                # if we only have one number left, lets call that the lot number
                number = alpha_split[0]
                # if len(alpha_split) == 1:
                #     number = alpha_split[0]
                # # if we had letters/numbers/letters/numbers, we will have len(alpha_split) > 1, so we cant decide
                # else:
                #     raise Exception(f'Tried splitting the Lot ({lot}/{cleaned}) by letters, but got more than 1 reminaing value: {alpha_split}')
                
                islot = True
            except Exception as e:
                raise e
    
    
    if islot:
        out_start = 'T'
        
        
    elif isticket:
        out_start = 'TKT'
        
        if len(cleaned_parts) > 1:
            cleaned = cleaned_parts[1]
        
        
        
        # split it on the numbers
        number_split = re.split(r'(\d+)', cleaned)
        # get rid of anything that is a space or empty line
        number_split = [i for i in number_split if not re.match(r'(\s)|(^$)',i)]
        
        
        # this is likely when we have len(cleaned_parts) > 1 
        if len(number_split) == 1:
            number = number_split[0]
        # for other instances we should end up here, getting the 2nd item in the number_split list
        elif len(number_split) == 2:
            number = number_split[1]
        else:
            
            print(f"Tried splitting the TKT ('{lot}' -> '{cleaned}') by numbers, but got not 2 results: {number_split}")
            return None
    

    # what do we do when len(number) > 3
    number = number.zfill(3)
    
    out = out_start + number
    
    print(f'{lot} --> {out}')
    return out
        
    



def insert_fablisting_to_live(fablisting, source=None):
    fablisting_df = further_fablisting_df_preperations(fablisting) # move this to get_model_estimate_hours_attached_to_fablisting_SQL
    fablisting_df['lotcleaned'] = fablisting_df['lot #'].apply(lotnumber_cleaner)
    
    table = 'fablisting'
    
    print_count_results(engine, schema='live', table=table, suffix_text='before truncating')
    
    
    if table_exists(engine, 'live', table):
        Session = sessionmaker(bind=engine)
        session = Session()
        session.execute(text(f'''TRUNCATE TABLE live.{table}'''))
        session.commit()
        session.close()
    
    
    # get count of table before insert --> should be 0
    print_count_results(engine, schema='live', table=table, suffix_text='before importing')
    # print_count_results(engine, schema='dbo', table=table, suffix_text='before merge proc')
    
    fablisting_df.to_sql(table, engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine, schema='live', table=table, suffix_text='after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call dbo.merge_fablisting(%s)", (source,))
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine, schema='dbo', table=table, suffix_text='after merge proc')
    print_count_results(engine, schema='live', table=table, suffix_text='after merge proc')
