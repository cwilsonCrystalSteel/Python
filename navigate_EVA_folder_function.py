# -*- coding: utf-8 -*-
"""
Created on Wed Feb 16 16:25:52 2022

@author: CWilson
"""
import os 
import glob
import pandas as pd

def get_df_of_all_lots_files_information():
    
    # this can still get you access to the mapped drive even if the mapped drive is unconnected
    prereq = '//192.168.50.9//Dropbox_(CSF)//'
    os.scandir(prereq)
    
    # this is the base dir from the X:\\
    base_dir = 'X:\\production control\\EVA REPORTS FOR THE DAY\\'
    
    # this is the base dir from the Universal Name Convention (UNC)
    # this seems to work even if it is disconnected!
    # to access this: cmd.exe -> 'net use' -> remote 
    base_dir = '//192.168.50.9//Dropbox_(CSF)//production control\\EVA REPORTS FOR THE DAY\\'
    
    
    
    
    try:
        years = os.listdir(base_dir)
    except:
        print('COULD NOT CONNECT TO THE X DRIVE / DROPBOX')
        exit()
        
        
        
        
    if 'desktop.ini' in years:
        years.remove('desktop.ini')
    
    big_ole_dict = {}
    df = pd.DataFrame(columns=['year','month','day','job','lot','shop','basename','destination','ez_dir'])
    
    jobs = {}
    
    for year in years:
        
        months = os.listdir(base_dir + year)
        
        big_ole_dict[year] = {}
        
        for month in months:
            
            days = os.listdir(base_dir + year + '\\' + month)
            
            big_ole_dict[year][month] = {}
            
            for day in days:
                
                this_day = os.listdir(base_dir + year + '\\' + month + '\\' + day)
                
                xls_files = glob.glob(base_dir + year + '\\' + month + '\\' + day + "\\*.xls")
                
                big_ole_dict[year][month][day] = []
                
                for xls_file in xls_files:
                    
                    basename = os.path.basename(xls_file)
                    
                    basename_components = basename.split('-')
                    
                    if len(basename_components) == 3:
                        # job # is the start of the filename
                        job_str = basename_components[0]
                        #
                        job_int = int(job_str)
                        # lot is the middle portion of the filename
                        lot = basename_components[1]
                        # chop off the '.xls' and only maintain the shop 
                        shop = basename_components[2][:-4]
                        
                        # skip the xls file if it does not have a valid name
                        if lot[0] != 'T':
                            continue
                    
                    else:
                        print('ERROR THE FILENAME IS INVALID')
                        ''' SEND ERROR THAT THE FILENAME IS INVALID
                        1) Emad
                        2) mildred tong
                        3) me
                        '''
                    print(day, month, year)
                    
                    ez_dir = 'X:\\' + xls_file[31: xls_file.find(basename)]
                    
                    df = df.append({'year':year, 'month':month, 'day':day, 'job':job_int,'lot':lot,'shop':shop,'basename':basename, 'destination':xls_file, 'ez_dir':ez_dir}, ignore_index=True)
                    big_ole_dict[year][month][day].append(basename)
                    
    return df