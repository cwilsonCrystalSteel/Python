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

def fitter_welder_stats_month(month_num=3, year=2025): 
    
    start_dt = datetime.date(year, month_num, 1)
    if month_num == 12:
        next_month = datetime.date(year + 1, 1, 1)
    else:
        next_month = datetime.date(year, month_num + 1, 1)
    end_dt = next_month - datetime.timedelta(days=1)
    
    # convert datetimes into string
    start_date = start_dt.strftime('%m/%d/%Y')
    end_date = end_dt.strftime('%m/%d/%Y')
    
    
    # get clock info for time period
    times_df = get_timesdf_from_vClocktimes(start_date, end_date)
    # get the basis of information
    basis = return_basis_new_direct_rules(times_df, include_terminated=True, remove_CSM_shop_b=False)
    # employee information
    ei = basis['Employee Information']
    
    
    
    
    
    direct = basis['Direct']
    indirect = basis['Indirect']
    notcounted = basis['Not Counted']
    missedhours = basis['Missed Hours']
    
    direct['direct'] = 'direct'
    indirect['direct'] = 'indirect'
    notcounted['direct'] = 'notcounted'
    missedhours['direct'] = 'missed'
    
    ''' we want to join all hours together'''
    hours_list = []
    for hours_type in [direct, indirect, notcounted, missedhours]:
        # make a copy of the df
        hours_type = hours_type.copy()
        # get only the select columns
        hours_type = hours_type[['Name','Location','Hours','direct']]
        # group by and sum the hours
        hours_type_grouped = hours_type.groupby(['Name','Location','direct']).sum()
        # bring it back to a df
        hours_type_grouped = hours_type_grouped.reset_index()
        # add to the list
        hours_list.append(hours_type_grouped)
    
    # join all from the lsit into one df
    hours_type_combined = pd.concat(hours_list)
    # pivot to get the different types of hours as columns
    hours_types_pivot = hours_type_combined.pivot_table(index=["Name", "Location"], columns="direct", values="Hours", aggfunc="sum").reset_index()
    # replace na with 0
    hours_types_pivot = hours_types_pivot.fillna(0)
    # calculate totla hours worked
    hours_types_pivot['total'] = hours_types_pivot[['direct','indirect','notcounted']].sum(axis=1)
    
    ''' Fablisting & Defects '''
    
    
    
    all_fitters_dict = {}
    all_welders_dict = {}
    fix_me_fitter_dict = {}
    fix_me_welder_dict = {}
    wrong_state_fitter_dict = {}
    wrong_state_welder_dict = {}
    employee_not_worked_fitter_dict = {}
    employee_not_worked_welder_dict = {}
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
        
        
        array_of_ids = fitters
        col_name = 'Fitter'
        # get only the fitter information retrieved form the dataframe
        fitter_data = return_sorted_and_ranked(df=df, ei=ei, array_of_ids=array_of_ids, 
                                               col_name=col_name, defect_log=defect_log, 
                                               state=state, start_date=start_date, end_date=end_date,
                                               hours_types_pivot=hours_types_pivot)
        
        array_of_ids = welders
        col_name = 'Welder'
        # get only the welder information retrieved form the dataframe
        welder_data = return_sorted_and_ranked(df=df, ei=ei, array_of_ids=array_of_ids, 
                                               col_name=col_name, defect_log=defect_log, 
                                               state=state, start_date=start_date, end_date=end_date,
                                               hours_types_pivot=hours_types_pivot)
        

            
        all_fitters_dict[state] = fitter_data['employees'].copy().reset_index(drop=True)
        all_welders_dict[state] = welder_data['employees'].copy().reset_index(drop=True)
        fix_me_fitter_dict[state] = fitter_data['df_to_fix']
        fix_me_welder_dict[state] = welder_data['df_to_fix']
        wrong_state_fitter_dict[state] = fitter_data['df_wrong_state']
        wrong_state_welder_dict[state] = welder_data['df_wrong_state']
        employee_not_worked_fitter_dict[state] = fitter_data['df_notworked']
        employee_not_worked_welder_dict[state] = welder_data['df_notworked']
            
        
    # combine
    all_both = pd.concat(list(all_fitters_dict.values()) +  list(all_welders_dict.values()), axis=0)
    all_both = all_both.sort_values(by=['Classification','Weight'], ascending=False)
    
    # join to the hours worked
    out_df = pd.merge(left=all_both, left_on=['Name','Location'],
                      right=hours_types_pivot, right_on=['Name','Location'],
                      how='left')
            
    # where to put the csv files
    directory = Path().home() / 'documents' / 'FitterWelderPerformanceCSVs'
    if not os.path.exists(directory):
        os.mkdir(directory)
    # create a timestamp so as not to override files
    file_timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # the time span of when the file was created for
    file_range = f"{start_dt.year}-{str(start_dt.month).zfill(2)}"
    # the end of the file name is the file range and creation timestamp 
    file_name_end = '_' + file_range + '_' + file_timestamp + '.csv'
    

    
    file_out = directory / ('all_both' + file_name_end)
    out_df.to_csv(file_out)

    return {'filepath':file_out, 
            'fix_fitters':fix_me_fitter_dict, 
            'fix_welders':fix_me_welder_dict,
            'wrong_state_fitters':wrong_state_fitter_dict,
            'wrong_state_welders':wrong_state_welder_dict,
            'employee_didnt_work_fitters':employee_not_worked_fitter_dict,
            'employee_didnt_work_welders':employee_not_worked_welder_dict}


