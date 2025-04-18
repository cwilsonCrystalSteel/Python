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


def preprocess_data(times_df, include_terminated=False):

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
    
    
    # split the cost code on a space or a slash
    df['Job #'] = df['Cost Code'].str.split(r'\s|\\').str[0]
    # 2023-09-28: changing this to the 1st 4 digits b/c apparantly thats what it is 
    df['Job #'] = df['Job #'].str[:4]
    
    # here is another way that will ensure we get all of the numbers - we need for when the job # is 6 digits
    df['Job # 2'] = df['Cost Code'].str.extract(r'^(\d+)')
    
    # get the shop site by taking the first 2 characters of the PRODUCTIVE tag from the EI dataframe
    df = df.join(ei.set_index('Name')['Productive'].astype(str).str[:2], on='Name')
    # rename that to location
    df = df.rename(columns={'Productive':'Location'})
    
    # reset the index of times_df
    df = df.reset_index(drop=True)
    
    
    return df, ei


def determine_absent(ei_subset, df):
    
    
    # get the employees that are in ei_shop but not in times_df
    absent = ei_subset[~ei_subset['Name'].isin(df['Name'])]
    # this is to prevent an error
    absent_copy = absent.copy()
    # create a shop name column based on the Productive column
    absent_copy.loc[:,'Shop'] = absent.loc[:,'Productive'].str[:2]
    # put it back to the original
    absent = absent_copy.copy()
    # drop the columns not needed
    absent = absent.drop(columns=['First','Last','Location','Shift'])
    
    return absent


def remove_csm_shop_b_employees(ei, df):
    ''' remove CSM shop B employees here '''
    # remove employee ids: 2001, 2015, 2029
    # get the names of employees 2001, 2015, 2029
    shop_b_employees = ei[ei['ID'].isin([2001,2015,2029,2242,2261,2007,2241])]
    # drop those names from times_df
    df = df[~df['Name'].isin(list(shop_b_employees['Name']))]
    
    return df

def determine_absent_hours_that_got_clocked(df):
    # get any clock items that correspond to missing hours / an abssence 
    absent_clocked = df[(df['Cost Code'] == '646 - Excused Absence') |
                        (df['Cost Code'] == '646 EXCUSED ABSENCE') |
                        (df['Cost Code'] == '647 - Unexcused Absence') |
                        (df['Cost Code'] == '648 - Unpaid time off request') |
                        (df['Cost Code'] == '649 - PTO (Office)') |
                        (df['Cost Code'] == '650 - PTO (Shop)') |
                        (df['Cost Code'].str.startswith('650 PTO')) |
                        (df['Cost Code'] == '653 - Bereavement') |
                        (df['Cost Code'] == '654 - Work Comp') |
                        (df['Cost Code'] == '656 - Emergency Paid Leave') |
                        (df['Cost Code'] == '657 - Union Sick Time') |
                        (df['Cost Code'] == '658 - Union Vacation') 
                        ]
                        
    



    return absent_clocked
    

def determine_how_many_hours_missed_out_on(absent, absent_clocked):
    ''' this will be a way to show how many potential hours were left out '''
    # get any of the records from absent_clocked that have more than zero hours
    missed_hours = absent_clocked[absent_clocked['Hours'] > 0]   
    # get only the select columns
    missed_hours = missed_hours[['Hours','Job Code','Name','Location']]

    ''' 
    2025-04-08
    I think this will have to be addressed! 
    we may need to update the sql merge_terminated 
    to try and figure out the date of termiantion
    '''
    # this is to make sure even if we have include_terminated=True, we dont count them as contributing 
    absent_missed = absent[~absent['terminated']]
    # append the absent dataframe columns of Name & Shop to the bottom
    missed_hours = pd.concat([missed_hours, absent_missed[['Name','Shop']]])
    # set the Location column to be coalesce(location, shop)
    missed_hours['Location'] = missed_hours['Location'].combine_first(missed_hours['Shop'])
    # get rid of the shop column now
    missed_hours = missed_hours.drop(columns='Shop')
    # when coming from the absent df, we will say they missed 8 hours
    missed_hours['Hours'] = missed_hours['Hours'].fillna(8)
    # when coming from the absent df, we will say the reason is they didnt clock in
    missed_hours['Job Code'] = missed_hours['Job Code'].fillna('Did Not Clock In')
    # change the name from job code to reason
    missed_hours = missed_hours.rename(columns={'Job Code':'Reason'})

    return missed_hours    
    

def return_information_on_clock_data(times_df, include_terminated=False, remove_CSM_shop_b=True):
    df, ei = preprocess_data(times_df, include_terminated)
        
    # remove employees without a productive/nonproductive grouping
    ei_shop = ei[~ei['Productive'].isna()]
    # get only people that have productive in their schedule group
    ei_shop = ei_shop[ei_shop['Productive'].str.contains('PRODUCTIVE')]
    
    # get the people who are not in the times_df, but in the subset of employees
    absent = determine_absent(ei_shop, df)
    
    if remove_CSM_shop_b:
        # get rid of shop B
        df = remove_csm_shop_b_employees(ei, df)
    
    
    # get items where length of the job # is 5 long
    # or the cost code says 250 RECEIVING {SHOP}
    # or the cost code says Transfers (for fed its job # 355)
    direct = df[(df['Job #'].str.len() == 5) | 
                (df['Cost Code'].str.contains('500 RECEIVING')) |
                (df['Cost Code'].str.contains('250 RECEIVING')) |
                (df['Cost Code'].str.contains('TRANSFERS'))].copy()
    
    # when working TRANSFERS, those hours are 100% direct, and count as 100% earned hours
    direct_as_earned = direct[direct['Cost Code'].str.contains('TRANSFERS')]
    # inidirect is whatver is not in the direct dataframe (in this instance)
    indirect = df.loc[~df.index.isin(direct.index)].copy()


    

    absent_clocked = determine_absent_hours_that_got_clocked(df)
    # now lets remove those absent hours that got clocked just to be sure 
    direct = direct.loc[~direct.index.isin(absent_clocked.index)].copy()    
    # and remove those from the indirect hours 
    indirect = indirect.loc[~indirect.index.isin(absent_clocked.index)].copy()   

    # get the df of how many hours missed & why
    missed_hours = determine_how_many_hours_missed_out_on(absent, absent_clocked)


    # label the direct dataframe as having direct
    direct['Is Direct'] = True
    # label the indirect dataframe as not having direct
    indirect['Is Direct'] = False
    

    
    return {'Absent':absent, 
            'Employee Information':ei_shop, 
            'Clocks Dataframe':times_df,
            'Direct':direct,
            'Indirect':indirect,
            'Missed Hours':missed_hours,
            'Direct as Earned Hours':direct_as_earned}    


def return_basis_new_direct_rules(times_df, include_terminated=False, productive_only=False, remove_CSM_shop_b=True):
    df, ei = preprocess_data(times_df, include_terminated)
        
    # remove employees without a productive/nonproductive grouping
    ei_shop = ei[~ei['Productive'].isna()]
    if productive_only:
        # get only people that have productive in their schedule group
        ei_shop = ei_shop[(ei_shop['Productive'].str.contains('PRODUCTIVE')) & 
                          (~ei_shop['Productive'].str.contains('NON PRODUCTIVE'))]
    
    
    df_productive = pd.merge(left=df,
                             right=ei_shop,
                             left_on='Name',
                             right_on='Name',
                             how='inner',
                             suffixes=('','_remove'))
    
    df = df_productive[[i for i in df_productive.columns if i in df.columns]].copy()
    
    
    
    # get the people who are not in the times_df, but in the subset of employees
    absent = determine_absent(ei_shop, df)
    
    # figure out the hours that correspond to being absent 
    absent_clocked = determine_absent_hours_that_got_clocked(df)

    # get the df of how many hours missed & why
    missed_hours = determine_how_many_hours_missed_out_on(absent, absent_clocked)
    
    if remove_CSM_shop_b:
        # get rid of shop B
        df = remove_csm_shop_b_employees(ei, df)
    
    
    ''' 2025-04-08 (or earlier)
    NEW DIRECT & INDIRECT RULES
    1) Does not matter if they are productive employees working non-productive codes
    Indirect: Quality Control
    Not Counted: Job #: 100, 175, 275, 525, 800, >100,000
    Direct:
    '''
    
    notcount = df[(df['Job #'] == '100') |
                  (df['Job #'] == '175') |
                  (df['Job #'] == '275') |
                  (df['Job #'] == '525') |
                  (df['Job #'] == '760') |
                  (df['Job #'] == '770') |
                  (df['Job #'] == '780') |
                  (df['Job #'] == '800') |
                  (df['Job # 2'].str.len() >=6) # when it is a 6 digit job like 930100
                  ].copy()
    
    indirect = df[(df['Job #'] == '200') |
                  (df['Cost Code'].str.contains('QUALITY CONTROL')) |
                  (df['Job #'] == '300') |
                  (df['Job #'] == '325') |
                  (df['Job #'] == '350') |
                  (df['Job #'] == '375') |
                  (df['Job #'] == '500') 
                  ].copy()
    
    # direct is what we have left 
    direct = df.loc[(~df.index.isin(notcount.index)) & 
                    (~df.index.isin(indirect.index)) & 
                    (~df.index.isin(absent_clocked.index))
                    ].copy()
    
        
    # when working TRANSFERS, those hours are 100% direct, and count as 100% earned hours
    direct_as_earned = direct[direct['Cost Code'].str.contains('TRANSFERS')]


    # label the direct dataframe as having direct
    direct['Is Direct'] = True
    # label the indirect dataframe as not having direct
    indirect['Is Direct'] = False
    

    
    return {'Absent':absent, 
            'Employee Information':ei_shop, 
            'Clocks Dataframe':times_df,
            'Direct':direct,
            'Indirect':indirect,
            'Not Counted':notcount,
            'Missed Hours':missed_hours,
            'Direct as Earned Hours':direct_as_earned}    