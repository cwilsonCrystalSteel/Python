# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 16:48:40 2025

@author: Netadmin
"""

import pandas as pd
import numpy as np
import datetime
from pathlib import Path
import os
from TimeClock.pullGroupHoursFromSQL import get_timesdf_from_vClocktimes
from TimeClock.functions_TimeclockForSpeedoDashboard import return_basis_new_direct_rules
from Grab_Defect_Log_Google_Sheet_Data import grab_defect_log
from Fitter_Welder_Stats_functions_v2 import clean_and_adjust_fab_listing_for_range
from Fitter_Welder_Stats_functions_v2 import return_sorted_and_ranked
from Fitter_Welder_Stats_functions_v2 import convert_weight_to_earned_hours
from Fitter_Welder_Stats_functions_v2 import get_employee_name_ID
from Fitter_Welder_Stats_functions_v2 import download_employee_group_hours
from Fitter_Welder_Stats_functions_v2 import get_employee_hours
from Fitter_Welder_Stats_functions_v2 import combine_multiple_all_both_csv_files_into_one_big_one



state = 'TN'
start_date = "02/02/2025"
end_date = "02/28/2025"
states = ['TN','MD','DE']


times_df = get_timesdf_from_vClocktimes(start_date, end_date)
basis = return_basis_new_direct_rules(times_df, include_terminated=True)
ei = basis['Employee Information']

start_dt = datetime.datetime.strptime(start_date, '%m/%d/%Y')
end_dt = datetime.datetime.strptime(end_date, '%m/%d/%Y')




# hours_df = clock_df
hours_df = times_df.copy()
direct = basis['Direct']
indirect = basis['Indirect']

direct['Hours'] = direct['Hours'].astype(np.float64)
indirect['Hours'] = indirect['Hours'].astype(np.float64)



for state in states:
    print(state)
    # get the defect log data for that time range
    # this way it only pulls the defect log information once per state
    defect_log = grab_defect_log(state, start_date, end_date)
    
    # get the dataframe of FabListing for that time range & state
    fablisting_cleaned = clean_and_adjust_fab_listing_for_range(state, 
                                                                start_date, 
                                                                end_date, 
                                                                earned_hours = 'best')
    df = fablisting_cleaned['Fab df'].copy()
    df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce')
    # df = df.rename(columns={'Earned Hours':'EVA Hours'})
    # it also gets the unique fitters and welders from the dataframe
    fitters = fablisting_cleaned['Fitter list']
    welders = fablisting_cleaned['Welder list']
    
    
    ''' IF EITHER OF THE FITTER_DATA OR WELDER_DATA THROWS AN ERROR REFER TO THE DEFINED FUNCTION '''
    ''' Or you can just run from line 72 down to these functions and look at the printed output '''
    # get only the fitter information retrieved form the dataframe
    fitter_data = return_sorted_and_ranked(df, ei, array_of_ids=fitters, 
                                           col_name="Fitter", defect_log=defect_log, 
                                           state=state, start_date=start_date, end_date=end_date)
    
    # get only the welder information retrieved form the dataframe
    welder_data = return_sorted_and_ranked(df, ei,array_of_ids= welders, 
                                           col_name="Welder", defect_log=defect_log, 
                                           state=state, start_date=start_date, end_date=end_date)
    
    fix_me_fitter = fitter_data['df_to_fix']
    fix_me_welder = welder_data['df_to_fix']
    
    # if it is the first state in the list, create a dataframe from the current dataframes
    if state == states[0]:
        all_fitters = fitter_data['employees'].copy().reset_index(drop=True)
        all_welders = welder_data['employees'].copy().reset_index(drop=True)
    # if it is not the first state, jsut append the data to the company wide dataframe
    else:
        all_fitters = pd.concat([all_fitters, fitter_data['employees']], ignore_index=True)
        all_welders = pd.concat([all_welders, welder_data['employees']], ignore_index=True)
        
    
        
# where to put the csv files
directory = Path().home() / 'documents' / 'FitterWelderPerformanceCSVs'
if not os.path.exists(directory):
    os.mkdir(directory)
# create a timestamp so as not to override files
file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
# the time span of when the file was created for
file_range = start_date.replace('/','-') + '_to_' + end_date.replace('/','-')
# the end of the file name is the file range and creation timestamp 
file_name_end = '_' + file_range + '_' + file_timestamp + '.csv'
# convert the fitter_data df to a csv
# fitter_data.to_csv(file_name_start + 'fitters' + file_name_end)
# # convert the welder_data df to a csv
# welder_data.to_csv(file_name_start + 'welders' + file_name_end)



all_fitters_grouped = get_employee_hours(all_fitters, direct, indirect)
all_fitters_grouped = all_fitters_grouped.round(3)
all_welders_grouped = get_employee_hours(all_welders, direct, indirect)
all_welders_grouped = all_welders_grouped.round(3)

# combine
all_both = pd.concat([all_fitters, all_welders], axis=0)
all_both = all_both.sort_values(by=['Classification','Weight'], ascending=False)
all_both.to_csv(directory / ('all_both' + file_name_end))


# after completing the loop, convert all_fitters to a csv
all_fitters_grouped.to_csv(directory / ('all_fitters' + file_name_end))
# after completing the loop, convert all_welders to a csv
all_welders_grouped.to_csv(directory / ('all_welders' + file_name_end))



