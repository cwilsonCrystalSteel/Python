# -*- coding: utf-8 -*-
"""
Created on Wed Jul 21 10:33:56 2021

@author: CWilson
"""

import pandas as pd
from Read_Group_hours_HTML import new_output_each_clock_entry_job_and_costcode
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
import datetime
import json
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours
code_changes = json.load(open("C:\\users\\cwilson\\documents\\python\\job_and_cost_code_changes.json"))



''' This is the main directory you will be working out of '''

base = 'C:\\users\\cwilson\\documents\\bonus\\'



''' 
1) Download the group hour data for all shops, make sure to include suspended
    and terminated employees when setting the filter!!! 
    - Must login with JTURNER account to have access to old data in Group Hours

2) Download a new employee information csv file and put it into the bonus folder
    - Must login with CODYW account 
    - Go to Tools -> Export -> Export Type = Employee Information -> Export Templates = Employee Locations

3) Name the files "MONTH YEAR.html" (ex. "July 2021.html") and 
    "MONTH YEAR Employee Information.csv" (ex. "July 2021 Employee Information.csv")

4) then simply change the values for the variables:
    -> start_date
    
5) Press the run arrow at the top menu bar
'''

start_date = '10/01/2021'



states = ['TN','MD','DE']
averages = pd.read_excel('c:\\downloads\\averages.xlsx')






start_dt = datetime.datetime.strptime(start_date,'%m/%d/%Y')
month = datetime.datetime.strftime(start_dt, '%B')
year = str(start_dt.year)

end_dt = datetime.datetime(start_dt.year, start_dt.month+1, 1) - datetime.timedelta(days=1)
end_date = end_dt.strftime('%m/%d/%Y')
# read the html file into a dataframe
times_df = new_output_each_clock_entry_job_and_costcode(base + month + ' ' + year + '.html')
# read the employee information csv file
ei = pd.read_csv(base + month + ' ' + year +' Employee Information.csv')
# rename shit in the employee infor df
ei = ei.rename(columns={'<NUMBER>':'ID',
                     '<FIRSTNAME>':'First',
                     '<LASTNAME>':'Last',
                     '<LOCATION>':'Location',
                     '<CLASS>':'Shift',
                     '<SCHEDULEGROUP>':'Productive'})
    
# combine the name fields to a combined single field    
ei['Name'] = ei['First'] + ' '+ ei['Last']  
# get rid of any duplicate names & keep the last entry with the newest ID number
ei = ei.loc[~ei['Name'].duplicated(keep='last')]
# remove employees without a productive/nonproductive grouping
ei_shop = ei[~ei['Productive'].isna()]
# get only people that have productive in their schedule group
ei_shop = ei_shop[ei_shop['Productive'].str.contains('PRODUCTIVE')]
ei_shop = ei_shop.drop(columns=['Location'])
ei_shop['Location'] = ei_shop['Productive'].str[:2]
# merge the location column based on the employee names
times_df = times_df.join(ei_shop.set_index('Name')['Location'], on='Name')

#%%
# define the states for the for loop
states = ['TN','DE','MD']

# states = ['TN']
#loop thru each state
for state in states:

    # get that states data
    hours = times_df[times_df['Location'] == state]
    # replace nocostcode with a dash
    hours.loc[:,'Cost Code'] = hours['Cost Code'].replace('no cost code', '-')
    # delete the stuff like pto
    hours = hours[~hours['Job Code'].isin(code_changes['Delete Job Codes'])]
    # create the indirect df starting with indirect cost codes
    indirect = hours[hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # remove those from the hours df now
    hours = hours[~hours['Cost Code'].isin(code_changes['Indirect Cost Codes'])]
    # append the indirect job codes
    indirect = indirect.append(hours[hours['Job Code'].isin(code_changes['Indirect Job Codes'])])
    # remove those from the hours df now
    hours = hours[~hours['Job Code'].isin(code_changes['Indirect Job Codes'])]
    # start with the things that appear to be indirect but are actually direct
    direct = hours[hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # remove those from the hours df now
    hours = hours[~hours['Job Code'].isin(code_changes['Direct Job Codes'])]
    # get the 3 digit items from hours df and move to indirect
    three_digit_indirects = hours[hours['Job #'].astype(int).astype(str).str.len() < 4]
    # append the 3digit jobs that are indirect into indirect df
    indirect = indirect.append(three_digit_indirects)
    # delete the 3digit jobs from hours
    hours = hours[~hours.index.isin(three_digit_indirects.index)]
    # append what is left of hours df to direct
    direct = direct.append(hours)
    
    ''' We need to move Cost Codes [300, 710, 720, 730] (Revisions & Work Tickets)
        from indirect to direct hours 
        Also need to sum up all of those hours and add that to the earned hours
        Because we want to assume 100% efficiency on tickets & revisions
    '''
    # define the cost codes that get swapped & then added to overall earned
    revisions_tickets_list = ['300 REVISIONS',
                              "710 CSF TICKETS",
                              "720 CSM TICKETS",
                              "730 FED TICKETS"]
    # get those cost code entries from the indirect df
    revisions_tickets_df = indirect[indirect['Cost Code'].isin(revisions_tickets_list)]
    # drop those rows from indirect
    indirect = indirect.drop(index=revisions_tickets_df.index)
    # append those rows to direct
    direct = direct.append(revisions_tickets_df)    
    # calculate the total number of hours for the revisions & tickets
    revisions_tickets_hours_grouped = revisions_tickets_df.groupby('Job #').sum()
    # rename the revisions & tickets stuff to earned hours
    revisions_tickets_hours_grouped = revisions_tickets_hours_grouped.rename(columns={'Hours':'Ticket/Rev Hours'})
    
    ''' This little snippet just removes the cost code for the following jobs:
            1626, 400 '''
    # get rid of the cost code associated with 1626
    sixteentwentysix = indirect[indirect['Job #'] == 1626]
    # set the cost code to be a  '-'
    sixteentwentysix.loc[:,'Cost Code'] = pd.Series(index=sixteentwentysix.index, data='-')
    # set the 1626 data back into place in the indirect df
    indirect.loc[sixteentwentysix.index] = sixteentwentysix
    # get rid of the cost code associated with 400
    fourhundred = indirect[indirect['Job #'] == 400]
    # set the cost code to be a  '-'
    fourhundred.loc[:,'Cost Code'] = pd.Series(index=fourhundred.index, data='-')
    # set the 1626 data back into place in the indirect df
    indirect.loc[fourhundred.index] = fourhundred    
    # Sort the direct & indirect dfs
    direct = direct.sort_values(['Name','Job #','Cost Code'])
    indirect = indirect.sort_values(['Name','Job #','Cost Code'])
    
    ''' grouping and sorting values '''
    # need to drop the Job # column bc it is being added up like a bunch of numbers
    direct = direct.rename(columns={'Job #':'Count of Clock-Ins'})
    indirect = indirect.rename(columns={'Job #':'Count of Clock-Ins'}) 
    
    direct_summary = direct.groupby(['Job Code']).agg({'Hours':'sum', 'Count of Clock-Ins': 'count'})
    # direct by job number
    # direct_summary = direct.groupby(['Job Code']).sum()
    direct_summary = direct_summary.sort_values('Hours', ascending=False)
    # direct by job number and cost code
    direct_summary_wcc = direct.groupby(['Job Code','Cost Code']).agg({'Hours':'sum', 'Count of Clock-Ins': 'count'})
    direct_summary_wcc = direct_summary_wcc.sort_values('Hours', ascending=False)
    # indirect by job number and cost code
    indirect_summary = indirect.groupby(['Job Code','Cost Code']).agg({'Hours':'sum', 'Count of Clock-Ins': 'count'})
    indirect_summary = indirect_summary.sort_values('Hours', ascending=False)

    # get the sheet name for fablisting    
    if state == 'TN':
        sheet = 'CSM QC Form'
    elif state == 'MD':
        sheet = 'FED QC Form'
    elif state == 'DE':
        sheet = 'CSF QC Form'    
    
    # grab the fablisting data for the month
    fl = grab_google_sheet(sheet, start_date, end_date)
    ''' THIS IS WHERE YOU NEED TO CHANGE TO CONVERT THE EARNED HOURS
        CHANGE how='old way' to how='model' '''
    fl = apply_model_hours(fl, how='old way', shop=sheet[:3])
    # grouping the fablisting data - summing everythign except hours per ton
    # taking the average for hours per ton so that it just returns the hours per ton of the job if the how='old way' for earned hours
    wt_summary = fl.groupby(['Job #']).agg({'Weight':'sum','Quantity':'sum','Hours per Ton':'mean','Earned Hours':'sum'})
    # calculate the tons from the weight in lbs
    wt_summary['Tons'] = wt_summary['Weight'] / 2000
    
    # this is to get the revisions & ticket work hours into the earned

    wt_summary['Ticket/Rev Hours'] = revisions_tickets_hours_grouped['Ticket/Rev Hours']
    # replace any nan values with zero
    wt_summary = wt_summary.fillna(0)    
    wt_summary['Total Earned Hours'] = wt_summary['Earned Hours'] + wt_summary['Ticket/Rev Hours']

    # sort by the earned hours - high to low
    wt_summary = wt_summary.sort_values(by='Total Earned Hours', ascending=False)
    # reorder the wt_summary columns
    wt_summary = wt_summary[['Weight','Tons','Quantity','Hours per Ton','Earned Hours','Ticket/Rev Hours','Total Earned Hours']]
    # this is the final summary series for the first page of the excel file
    f_summary = pd.Series(index=['Indirect','Direct','Earned'],
                          data=[indirect['Hours'].sum(),
                                direct['Hours'].sum(),
                                wt_summary['Total Earned Hours'].sum()])
    # calculate the efficiency
    f_summary['Efficiency'] = f_summary['Earned'] / f_summary['Direct']
    
    f_summary['Ticket/Rev Work'] = revisions_tickets_hours_grouped['Ticket/Rev Hours'].sum()
    
    # set the name to be the state & month
    f_summary = f_summary.rename(state + ' ' + month)
    
    
    # write to an excel file 
    with pd.ExcelWriter(base + state + ' ' + month + ' 2021 Proof.xlsx') as writer:
        f_summary.to_excel(writer, sheet_name='Summary')
        wt_summary.to_excel(writer, sheet_name='Job Summary')
        direct_summary.to_excel(writer, sheet_name='Direct Hrs Summary')
        indirect_summary.to_excel(writer, sheet_name='Indirect Hrs Summary')
        direct_summary_wcc.to_excel(writer, sheet_name='Direct Hrs Summary 2')
        direct.to_excel(writer, sheet_name='Direct Hours', index=False)
        indirect.to_excel(writer, sheet_name='Indirect Hours', index=False)
        fl.to_excel(writer, sheet_name='Fablisting', index=False)
        
    
     
    print('\n\n\n')
    print('-'*50)
    if f_summary['Efficiency'] >= 0.85:
        print(state + ' Has hit the efficiency bonus for ' + month)
        print('-'*50)











