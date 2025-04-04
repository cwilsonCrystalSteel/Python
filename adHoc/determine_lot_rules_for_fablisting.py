# -*- coding: utf-8 -*-
"""
Created on Fri Apr  4 11:31:15 2025

@author: Netadmin
"""
from pathlib import Path
import os
import re
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
import pandas as pd

file  = Path(r'c:/users/netadmin/downloads/fed lot contains k.txt')
print(os.path.exists(file))

with open(file) as f:
    lines = [line.rstrip() for line in f]
    

tkt_strings = [re.split(r'(\d+)', re.sub(r'[ -]', '',i.upper()))[0] for i in lines]
unique_tkt_strings = list(set(tkt_strings))


fablisting = grab_google_sheet('FED QC Form', '01/01/2023','04/04/2025')
fablisting = pd.concat([fablisting, grab_google_sheet('CSM QC Form','01/01/2023','04/04/2025')])
fablisting = pd.concat([fablisting, grab_google_sheet('CSF QC Form','01/01/2023','04/04/2025')])


fed_all = list(set(fablisting['Lot #']))
fed_all = [i for i in fed_all if not 'ph' in i.lower()]
fed_all = [re.split(r'(\d+)', re.sub(r'[ -#-]', '', i.upper())) for i in fed_all]

fed_all0 = list(set([i[0] for i in fed_all]))
fed_all1 = list(set([i[1] for i in fed_all if len(i) > 1]))


x = fablisting[(fablisting['Lot #'].str.upper().str.contains('ITEM')) |
               (fablisting['Lot #'].str.upper().str.contains('RELEASE')) |
               (fablisting['Lot #'].str.upper().str.contains('LOTITEM'))]
#%%

def lotnumber_cleaner(lot):
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
                if len(alpha_split) == 1:
                    number = alpha_split[0]
                # if we had letters/numbers/letters/numbers, we will have len(alpha_split) > 1, so we cant decide
                else:
                    raise Exception(f'Tried splitting the Lot ({lot}/{cleaned}) by letters, but got more than 1 reminaing value: {alpha_split}')
                
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
        
    
for lot in list(fablisting['Lot #']):
    
    lot_out = lotnumber_cleaner(lot)
    
fablisting['lotcleaned'] = fablisting['Lot #'].apply(lotnumber_cleaner)


all_cleaned = list(set(fablisting['lotcleaned']))
all_cleaned_df = fablisting[['Lot #','lotcleaned']].drop_duplicates().sort_values(by='lotcleaned', ascending=False)

# fed_all = [i for i in fed_all if not 'ph' in i.lower()]
# fed_all = [re.split(r'(\d+)', re.sub(r'[ -#-]', '', i.upper())) for i in fed_all]
