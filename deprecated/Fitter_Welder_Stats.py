# -*- coding: utf-8 -*-
"""
Created on Tue Apr 20 09:18:43 2021

@author: CWilson
"""

import sys
sys.path.append('c://users//cwilson//documents//python//Weekly Shop Hours Project//')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\TimeClock')
from pullGroupHoursFromSQL import get_date_range_timesdf_controller
from functions_TimeclockForSpeedoDashboard import return_basis_new_direct_rules
# from TimeClock_Tools_Employee_Department import download_most_current_employee_department_csv
# from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Grab_Defect_Log_Google_Sheet_Data import grab_defect_log
import pandas as pd
import numpy as np
import datetime
from Fitter_Welder_Stats_functions import clean_and_adjust_fab_listing_for_range
from Fitter_Welder_Stats_functions import return_sorted_and_ranked
from Fitter_Welder_Stats_functions import convert_weight_to_earned_hours
from Fitter_Welder_Stats_functions import get_employee_name_ID
from Fitter_Welder_Stats_functions import download_employee_group_hours
from Fitter_Welder_Stats_functions import get_employee_hours
from Fitter_Welder_Stats_functions import combine_multiple_all_both_csv_files_into_one_big_one
from Gather_data_for_timeclock_based_email_reports import get_information_for_clock_based_email_reports
from Gather_data_for_timeclock_based_email_reports import skip_timeclock_automated_retrieval
from Gather_data_for_timeclock_based_email_reports import get_ei_csv_downloaded



state = 'TN'
start_date = "01/01/2024"
end_date = "12/30/2024"
states = ['TN','MD','DE']



# ei = get_employee_name_ID()
# hours_df = download_employee_group_hours(start_date, end_date)


# basis = get_information_for_clock_based_email_reports(start_date, end_date, exclude_terminated=False)


times_df = get_date_range_timesdf_controller(start_date, end_date)
basis = return_basis_new_direct_rules(times_df, include_terminated=True)
ei = basis['Employee Information']
# ei = get_ei_csv_downloaded(exclude_terminated=False)

start_dt = datetime.datetime.strptime(start_date, '%m/%d/%Y')
end_dt = datetime.datetime.strptime(end_date, '%m/%d/%Y')


# clock_df, direct, indirect = None, None, None


# for num_day in range(0, (end_dt-start_dt).days + 1):
# # for num_day in range(0,7):
    
#     day_str = (start_dt + datetime.timedelta(days=num_day)).strftime('%m/%d/%Y')
    
#     try:
#         this_day = get_information_for_clock_based_email_reports(day_str, day_str, exclude_terminated=False, ei=ei)
        
#         if this_day is None:
#             continue
#     except:
#         continue
    

    
#     # should not hit this part if this_day is none
#     loop_clock_df = this_day['Clocks Dataframe']
#     loop_direct_df = this_day['Direct']
#     loop_indirect_df = this_day['Indirect']
    
    
#     #iteratively add each days data to a main dataframe
#     if clock_df is None:
#         clock_df = loop_clock_df
#     else:
#         clock_df = clock_df.append(loop_clock_df, ignore_index=True)
#     if direct is None:
#         direct = loop_direct_df
#     else:
#         direct = direct.append(loop_direct_df, ignore_index=True)
#     if indirect is None:
#         indirect = loop_indirect_df
#     else:
#         indirect = indirect.append(loop_indirect_df, ignore_index=True)        
    



# hours_df = clock_df
hours_df = times_df.copy()
direct = basis['Direct']
indirect = basis['Indirect']




for state in states:
    print(state)
    # get the defect log data for that time range
    # this way it only pulls the defect log information once per state
    defect_log = grab_defect_log(state, start_date, end_date)
    
    # get the dataframe of FabListing for that time range & state
    fablisting_cleaned = clean_and_adjust_fab_listing_for_range(state, 
                                                                start_date, 
                                                                end_date, 
                                                                earned_hours = 'model')
                                                                # earned_hours='old way')
    df = fablisting_cleaned['Fab df']
    # df = df.rename(columns={'Earned Hours':'EVA Hours'})
    # it also gets the unique fitters and welders from the dataframe
    fitters = fablisting_cleaned['Fitter list']
    welders = fablisting_cleaned['Welder list']
    
    
    ''' IF EITHER OF THE FITTER_DATA OR WELDER_DATA THROWS AN ERROR REFER TO THE DEFINED FUNCTION '''
    ''' Or you can just run from line 72 down to these functions and look at the printed output '''
    # get only the fitter information retrieved form the dataframe
    fitter_data = return_sorted_and_ranked(df, ei, fitters, "Fitter", defect_log, state, start_date, end_date)
    # get only the welder information retrieved form the dataframe
    welder_data = return_sorted_and_ranked(df, ei, welders, "Welder", defect_log, state, start_date, end_date)
    
    # fitter_data = convert_weight_to_earned_hours(state, fitter_data, drop_job_weights=True)
    # welder_data = convert_weight_to_earned_hours(state, welder_data, drop_job_weights=True)
    
    
    
    
    # if it is the first state in the list, create a dataframe from the current dataframes
    if state == states[0]:
        all_fitters = fitter_data.copy().reset_index(drop=True)
        all_welders = welder_data.copy().reset_index(drop=True)
    # if it is not the first state, jsut append the data to the company wide dataframe
    else:
        all_fitters = pd.concat([all_fitters, fitter_data], ignore_index=True)
        all_welders = pd.concat([all_welders, welder_data], ignore_index=True)
        
    
        
# where to put the csv files
directory = 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//'
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
all_both.to_csv(directory + 'all_both' + file_name_end)


# after completing the loop, convert all_fitters to a csv
all_fitters_grouped.to_csv(directory + 'all_fitters' + file_name_end)
# after completing the loop, convert all_welders to a csv
all_welders_grouped.to_csv(directory + 'all_welders' + file_name_end)

print(directory + 'all_both' + file_name_end)



# all_both_csvs = ['c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_12-01-2022_to_12-31-2022_2023-01-10-18-15-56.csv',
#                 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_11-01-2022_to_11-30-2022_2023-01-11-10-52-01.csv',
#                 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_10-01-2022_to_10-31-2022_2023-01-12-10-46-09.csv',
#                 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_09-01-2022_to_09-30-2022_2022-10-14-12-01-38.csv',
#                 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_08-01-2022_to_08-31-2022_2022-10-14-08-17-24.csv',
#                 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_07-01-2022_to_07-31-2022_2022-10-13-18-05-02.csv']

# combine_multiple_all_both_csv_files_into_one_big_one(all_both_csvs, 'c://users//cwilson//documents//Fitter_Welder_Performance_CSVs//all_both_6month_July012022_December312022.csv')

