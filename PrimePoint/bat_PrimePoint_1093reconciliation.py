# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 16:49:58 2026

@author: Netadmin
"""

from PrimePoint.PrimePointNavigation import PrimePointEZ
from PrimePoint.JournalEntryMacroOperation import run_macro
import datetime
from pathlib import Path
import os
import shutil
import glob
import re


today = datetime.datetime.now()
friday = today
while friday.weekday() != 4:
    friday = friday + datetime.timedelta(days=1)
    
friday_str = friday.strftime('%Y-%m-%d')
friday_str_dir = friday.strftime('%Y.%m.%d')
    
attempts = 0

while attempts < 4:
    ppe = PrimePointEZ(friday_str)
    ppe.get_filepaths_wrapper()
    paths = ppe.filepaths_to_move
    ppe.kill()
    
    attempts += 1
    
    if len(paths):
        break 

if not len(paths):
    raise Exception ('COULD NOT GET THE PRIMEPOINT FILES TO MOVE')


##### create the folder in g-drive

''' NNED TO HAVE THE DRIVE AS A SHARED DRIVE AND NOT A SHARED-WITH-ME-FOLDER '''
# I AM GOING TO TRY THIS ONE FIRST G:\My Drive\2026 UNTIL I HAVE ACCESS TO IT IN A SHAREDRIVE
# possible_dir = [r"C:\Users\Netadmin\Downloads\test-g-drive", r'G:\Shared drives\Payroll\Primepoint\Payroll Imports\1093 (taxes) Upload']
possible_dir = [r'G:\Shared drives\Payroll\Primepoint\Payroll Imports\1093 (taxes) Upload']

for ii in possible_dir:
    
    base_dir = Path(ii)
    
    if os.path.exists(base_dir):
        print(f'Using the drive: "{base_dir}"')
        break
        
    else:
        continue

if not os.path.exists(base_dir):
    raise Exception('No directory could be determined')




####### define the dirs

this_week_dir = base_dir / str(friday.year) / friday_str_dir
last_week_friday = friday - datetime.timedelta(days=7)
last_week_friday_str_dir = last_week_friday.strftime('%Y.%m.%d')
last_week_dir = base_dir / str(last_week_friday.year) / last_week_friday_str_dir

if not os.path.exists(this_week_dir):
    os.makedirs(this_week_dir)
    print(f'Making directory: {this_week_dir}')


######## create subfolders
data_input_files = 'Data Input Files'
to_review_files = 'To Review'
dirs_to_create = [data_input_files,'Import','Upload files',to_review_files]

for ii in dirs_to_create:
    to_make = this_week_dir / ii
    if not os.path.exists(to_make):
        
        os.makedirs(to_make)
        print(f"Making directory: {to_make}")
    
    
    
    
####### Get the macro file to copy over    
    
# WE SHOULD GET THIS FROM THE LAST WEEK's DIRECTORY
last_week_macro_file = list(last_week_dir.glob('*.xlsm'))
# we dont find a macro file in last week's dir
if len(last_week_macro_file) == 0:    
    macro_file = Path(r'./PrimePoint/Macro File v2 - TEMPLATE.xlsm')
    print(f'Using default macro file from code repository: {macro_file}')
elif len(last_week_macro_file) == 1:
    macro_file = last_week_macro_file[0]
    print(f"Using last week's macro file: {macro_file}")
else:
    # if there are multiple macro files, try to get the newest based on its name v2 -> v1 -> etc.
    last_week_macro_file.sort(reverse=True)
    macro_file = last_week_macro_file[0]
    print(f"Using one of last week's {len(last_week_macro_file)} macro files (MULTIPLE FOUND): {macro_file}")


# copy the macro file over 
new_macro_file_template = this_week_dir / macro_file.name
shutil.copy(macro_file, new_macro_file_template)
    



####### place the files from PrimePoint paths into the correct location
input_files_dir = this_week_dir / data_input_files
for ii in paths:
    ii = Path(ii)
    new_dir = input_files_dir / ii.name
    shutil.copy(ii, new_dir)


# trigger the macros with appropriate state files

# get all the files that are available to import // pulled from the web
input_files = [f.name for f in input_files_dir.iterdir() if f.is_file()]
# figure out which states are present dynamically by first 2 cahr of all files
states = set([i[:2] for i in input_files])


# for each state do macro file 
for state in states:
    
    state_macro_file_name = f'Taxes {state} {friday_str_dir}{new_macro_file_template.suffix}'
    # make the full path 
    state_macro_file = this_week_dir / to_review_files / state_macro_file_name
    
    # get the bs file for that state
    bs_file = [i for i in input_files if state in i and 'BS' in i]
    # get the tax report file for that state
    tax_report_file = [i for i in input_files if state in i and 'Tax Report' in i]
    
    
    # just in case we have multiples try to get newest by timestamp at end of filename 
    bs_file.sort(reverse=True)
    tax_report_file.sort(reverse=True)
    
    bs_filepath = input_files_dir / bs_file[0]
    tax_report_filepath = input_files_dir / tax_report_file[0]

    run_macro(new_macro_file_template, state_macro_file, bs_filepath, tax_report_filepath)


# dun 
