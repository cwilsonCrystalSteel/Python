# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 16:49:58 2026

@author: Netadmin
"""

from PrimePoint.PrimePointNavigation import PrimePointEZ
import datetime

today = datetime.datetime.now()
friday = today
while friday.weekday() != 4:
    friday = friday - datetime.timedelta(days=1)
    
friday_str = friday.strftime('%Y-%m-%d')
    
ppe = PrimePointEZ(friday_str)
ppe.get_filepaths_wrapper()
paths = ppe.filepaths_to_move
ppe.kill()



''' NNED TO HAVE THE DRIVE AS A SHARED DRIVE AND NOT A SHARED-WITH-ME-FOLDER '''


# check access to google drive
def determine_directory_path():
    # possible_dir = ['Y:/','X:/','\\\\192.168.50.9\\Dropbox_(CSF)']
    possible_dir = [r'G:\Shared drives']
    for ii in possible_dir:
        base_dir = Path(ii) / 'production' / 'EVA Reports' /'EVA REPORTS FOR THE DAY'
        
        if os.path.exists(base_dir):
            print(f'Using the drive: "{base_dir}"')
            break
            
        else:
            continue
    
    # we did not find a match at this point!
    if not os.path.exists(Path(ii)):
        insertError('EVADropbox - determine_directory_path', 'could not reach any directory!')
        raise Exception('Could not determine directory for Dropbox EVA files.')
    
    return base_dir



# create the folder in g-drive
# create subfolders
# plaec the files from paths into the correct location
# move the macro file 
# trigger the macros with appropriate state files
# dun 
