# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 11:14:33 2024

@author: CWilson
"""


import pandas as pd
from TimeClock.pullEmployeeInformationFromSQL import return_sql_ei
from TimeClock.pullJobCostCodesFromSQL import return_sql_jobcostcodes




def replace_ids_with_columns(times_df, ei=None):
    """
    This will replace the employeeidnumber (CrystalSteel Employee ID from Timeclock)
        with Employee Name
    This will repalce the job_costcode_di (internal Database table id) with
        the job code, cost code

    Parameters
    ----------
    times_df : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    
    jccSQL = return_sql_jobcostcodes()
    if ei is None:
        ei = return_sql_ei()
    
    # if using get_timesdf_from_vClocktimes
    if 'jobcode' in times_df.columns and 'costcode' in times_df.columns:
        df = times_df.drop(columns=['targetdate','source'])
    # if using get_date_range_timesdf_controller
    else:
        df = pd.merge(left=times_df, right=jccSQL, how='left', left_on='job_costcode_id', right_on='id')
        
    df = pd.merge(left=df, right=ei, how='left', left_on='employeeidnumber', right_on='employeeidnumber')
    
    return df






def return_information_on_clock_data(times_df, include_terminated=False):
    
    ei = return_sql_ei()
    df = replace_ids_with_columns(times_df, ei=ei)
    
    # only what we need
    df = df[['hours','timein','timeout','jobcode','costcode','fullname']]
    # rename according to legacy code
    df = df.rename(columns = {'hours':'Hours',
                              'timein':'Time In',
                              'timeout':'Time Out',
                              'jobcode':'Job Code',
                              'costcode':'Cost Code',
                              'fullname':'Name'})
    
    # this is to line up with legacy code
    ei = ei.rename(columns={'employeeidnumber':'ID',
                     'firstname':'First',
                     'lastname':'Last',
                     'location':'Location',
                     'classshift':'Shift',
                     'schedulegroup':'Productive',
                     'department':'Department',
                     'fullname':'Name'})    
        
        
    if not include_terminated:
        # get active employees only for this 
        ei = ei[ei['terminated'] == False]
    # get rid of any duplicate names & keep the last entry with the newest ID number
    ei = ei.loc[~ei['Name'].duplicated(keep='last')]
    # remove employees without a productive/nonproductive grouping
    ei_shop = ei[~ei['Productive'].isna()]
    # get only people that have productive in their schedule group
    ei_shop = ei_shop[ei_shop['Productive'].str.contains('PRODUCTIVE')]
    

    
    # get the employees that are in ei_shop but not in times_df
    absent = ei_shop[~ei_shop['Name'].isin(df['Name'])]
    # this is to prevent an error
    absent_copy = absent.copy()
    # create a shop name column based on the Productive column
    absent_copy.loc[:,'Shop'] = absent.loc[:,'Productive'].str[:2]
    # put it back to the original
    absent = absent_copy.copy()
    # drop the columns not needed
    absent = absent.drop(columns=['First','Last','Location','Shift'])
    # reset the index of times_df
    df = df.reset_index(drop=True)
    
    #new way of getting direct / indirect
    # 3 digit codes are indirect - except for recieving
    # 5 digit codes are direct - except for CAPEX (idk how to tell if it is CAPEX)
    
    # split the cost code on a space or a slash (idk why i have to do 4 slashes)
    df['Job #'] = df['Cost Code'].str.split('\s|\\\\').str[0]
    # get the shop site by taking the first 2 characters of the PRODUCTIVE tag from the EI dataframe
    df = df.join(ei.set_index('Name')['Productive'].astype(str).str[:2], on='Name')
    # rename that to location
    df = df.rename(columns={'Productive':'Location'})
    
    ''' remove CSM shop B employees here '''
    # remove employee ids: 2001, 2015, 2029
    # get the names of employees 2001, 2015, 2029
    shop_b_employees = ei[ei['ID'].isin([2001,2015,2029,2242,2261,2007,2241])]
    # drop those names from times_df
    df = df[~df['Name'].isin(list(shop_b_employees['Name']))]
    ''' end of removing shop B employees '''
    
    # get items where length of the job # is 5 or the cost code says recieving (or I could do job #  = 250)
    direct = df[(df['Job #'].str.len() == 5) | (df['Cost Code'].str.contains('RECEIVING'))].copy()
    # inidirect is whatver is not in the direct dataframe (in this instance)
    indirect = df.loc[~df.index.isin(direct.index)].copy()
    # now to convert the job number to the actual job number - only need to remove the first digit from the 5 digit jobs
    # direct['Job #'] = direct['Job #'].str[-4:]
    # 2023-09-28: changing this to the 1st 4 digits b/c apparantly thats what it is 
    direct['Job #'] = direct['Job #'].str[:4]
    
       


    direct['Is Direct'] = True
    indirect['Is Direct'] = False
    

    
    return {'Absent':absent, 
            'Employee Information':ei_shop, 
            'Clocks Dataframe':times_df,
            'Direct':direct,
            'Indirect':indirect}    