# -*- coding: utf-8 -*-
"""
Created on Wed Apr 28 13:11:46 2021

@author: CWilson
"""
#%% the initial fireup of the file & getting data from TIMECLOCK

import pandas as pd
import os
from pathlib import Path
import datetime
import numpy as np
import copy
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from Lots_schedule_calendar.email_setup import get_email_service
import base64

# get today as a datetime
# today = datetime.datetime.now()
# # get yesterday as a datetime  
# yesterday = today - datetime.timedelta(days=1)
# # convert yesterday to a string format for my functions
# yesterday_str = yesterday.strftime("%m/%d/%Y")
# # get the data from TIMECLOCK then do basic transformations to it
# basis = get_information_for_clock_based_email_reports(yesterday_str, yesterday_str, exclude_terminated=True)



#%% This section is for the cleaning/processing functions

def groupby_unique_job_cost_code(raw_clock):
    # determine the unique names in the dataframe
    names = pd.unique(raw_clock['Name'])
    ''' This will summarize each unique Job-CostCode the employee worked on '''
    # create an empty list that will become a list of lists
    clock_details_df = []
    # iterate thru each employee
    for name in names:
        # get only the portion of the dataframe with their name
        chunk = raw_clock[raw_clock['Name'] == name]
        # get the unique job & costcode combinations
        job_costcodes = pd.unique(chunk['Job CostCode'])
        # get the state
        location = chunk.iloc[0]['Location']
        # iterate thru each unique job + cost code combinations
        for jcc in job_costcodes:
            # get the smaller chunk of the dataframe relating to that jcc
            jcc_df = chunk[chunk['Job CostCode'] == jcc]
            # get the job name
            job_number = jcc_df.iloc[0]['Job #']
            # get the job number
            job_code = jcc_df.iloc[0]['Job Code']
            # get the cost code
            costcode = jcc_df.iloc[0]['Cost Code']
            # sum up the hours for that jcc
            jcc_hours = jcc_df.sum()['Hours']
            # start = jcc_df.iloc[0]['Start'].time().strftime("%I:%M %p")
            # end = jcc_df.iloc[0]['End'].time().strftime("%I:%M %p")
            # append the name, job, costcode, jcc, and hours to the big list that will become a df later
            # clock_details_df.append([name, location, jobname, job, costcode, jcc, jcc_hours, start, end])
            clock_details_df = pd.concat([clock_details_df, 
                                          [name, location, job_number, job_code, costcode, jcc, jcc_hours]])
            # clock_details_df.append([name, location, job_number, job_code, costcode, jcc, jcc_hours])
    
    
    # convert the list of lists to a dataframe
    # clock_details_df = pd.DataFrame(data=clock_details_df, columns=['Name','Location','Job','Job #','Cost Code','JCC','Hours', 'Start', 'End'])
    clock_details_df = pd.DataFrame(data=clock_details_df, columns=['Name','Location','Job #','Job Code','Cost Code','JCC','Hours'])
    # get boolean series - true if length of job# is 4
    direct = clock_details_df['Job #'].astype(str).apply(len) == 4
    # rename the series
    direct = direct.rename('Is Direct')
    # add the column to the dataframe
    clock_details_df = clock_details_df.join(direct)
    
    return clock_details_df



def summarize_by_direct_indirect(grouped_clock):
    # determine the unique names in the dataframe
    names = pd.unique(grouped_clock['Name'])
    # create an empty list that will keep the summary data
    clock_summary_list = []
    # go thru each employee in the df
    for name in names:
        # get only their portion of the df
        chunk = grouped_clock[grouped_clock['Name'] == name]
        # get the portion that is direct hours
        direct_chunk = chunk[chunk['Is Direct']]
        # get the indirect portion
        indirect_chunk = chunk[~chunk['Is Direct']]
        # grab their location from the df
        location = chunk.iloc[0]['Location']
        # get the jobs only if there is a cost code with the word LOT in it 
        lots_only = direct_chunk[direct_chunk['Cost Code'].str.contains('LOT')]
        # count the number of lots 
        lots_count = lots_only.shape[0]
        # count the number of unique jobs
        jobs_count = pd.unique(chunk['Job #']).shape[0]
        # sum up the direct hours
        direct_hours = direct_chunk.sum()['Hours']
        # sum up the indirect hours
        indirect_hours = indirect_chunk.sum()['Hours']
        # append all the data to a list 
        clock_summary_list.append([name, location, jobs_count, lots_count, direct_hours, indirect_hours])
       
    # create the dataframe from the list
    clock_summary_df = pd.DataFrame(data=clock_summary_list, columns=['Name','Location','# Jobs','# Lots', 'Direct','Indirect'])
    # calculate the total hours
    clock_summary_df['Total'] = clock_summary_df['Direct'] + clock_summary_df['Indirect']
    # calculate the percentage of direct hours
    clock_summary_df['% Direct'] = np.round(clock_summary_df['Direct'] / clock_summary_df['Total'], 4)
    # get rid of anytime there is a clock with 0 hours in it for some reason
    clock_summary_df = clock_summary_df[~clock_summary_df['% Direct'].isna()]
    return clock_summary_df




def return_output_dictionary(filtered_summary_df, detail_df):
    # get the employee information from the clock_details_df if they are in sub_80
    small_df = detail_df[detail_df['Name'].isin(filtered_summary_df['Name'])]
    # get the states from the dataframe 
    states = pd.unique(filtered_summary_df['Location'])
    # create an empty dict that gets a detail & summary df for each state
    output_dict = {}
    # iterate thru each state in the small_df
    for state in states:
        # get the state specific dataframe
        summary_of_employees = filtered_summary_df[filtered_summary_df['Location'] == state]
        # we don't need to see the # Jobs or # Lots clocked in to 
        # summary_of_employees = summary_of_employees.drop(columns=['# Jobs','# Lots'])
        # initialize the dict within the state-key
        output_dict[state] = {}
        # the summary key gets the state summary
        output_dict[state]['Summary'] = summary_of_employees
        output_dict[state]['Detail'] = {}
        for employee in pd.unique(summary_of_employees['Name']):
        
            persons_df = small_df[small_df['Name'] == employee]
    
            output_dict[state]['Detail'][employee] = persons_df
            
    # put 2 empty lines after each employee in the detail_df of each state for HTML table reasons
    for state in states:
        deets = output_dict[state]['Detail']
        big_deet = pd.DataFrame()
        for df in deets.values():
            big_deet = pd.concat([big_deet, df.reset_index(drop=True)])
            # adds 2 empty rows after data
            big_deet.loc[big_deet.shape[0]] = pd.Series(dtype=int)
            big_deet.loc[big_deet.shape[0]] = pd.Series(dtype=int)
            # big_deet = big_deet.append(df, ignore_index=True)
            # big_deet = pd.concat([big_deet, pd.Series(dtype=int)], axis=1)
            # big_deet = big_deet.append(pd.Series(dtype=int), ignore_index=True)
            # big_deet = big_deet.append(pd.Series(dtype=int), ignore_index=True)
        big_deet = big_deet.drop(index=big_deet.index[-2:])
        output_dict[state]['Detail'] = big_deet    


    return output_dict



def return_pretty_string_format_of_df_dicts(dictionary):
    # grab the details dict for the state
    big_deet = dictionary['Detail'].copy()
    # get rid of the job # column and the Is Direct column
    big_deet = big_deet.drop(columns=['Job #', 'Is Direct'])   
    # convert hours to have 2 decimal places
    big_deet['Hours'] = pd.Series(["{0:.2f}".format(val) for val in big_deet['Hours']], index=big_deet.index)
    # replace string nan values with empty strings
    big_deet = big_deet.replace('nan', '')
    # replace np.nan values with empty strings
    big_deet = big_deet.replace(np.nan, '')
    # get rid of the location column
    # big_deet = big_deet.drop(columns=['Location', 'Start','End'])
    big_deet = big_deet.drop(columns=['Location'])
    # reset the index of the dataframe
    big_deet = big_deet.reset_index(drop=True)
    # create empty mapping dict
    new_cols = {}
    # iterate thru each column in the df
    for col in big_deet.columns:
        # give each column some underscores before and after it
        new_cols[col] = '-----' + col + '-----'
    # rename the column headers
    big_deet = big_deet.rename(columns=new_cols)
    # convert the dataframe to a string format table
    a = big_deet.to_html(col_space=100, index=False, justify='center')
    # replace the dataframe in the dictionary dict with the viewer friendly table
    dictionary['Detail'] = a
    
    
    # grab the df from the dict
    summary = dictionary['Summary'].copy()
    # get rid of location
    summary = summary.drop(columns=['Location'])
    # sort the table based on the percentage of direct hours worked
    summary = summary.sort_values(by='% Direct')
    # reset the index of the dataframe
    summary = summary.reset_index(drop=True)
    # conver the % direct column to a 2 decimal format with a percent sign
    summary['% Direct'] = pd.Series(["{:.2f}%".format(val * 100) for val in summary['% Direct']], index=summary.index)    
    # create empty mapping dict
    new_cols = {}
    # iterate thru each column in the df
    for col in summary.columns:
        # give each column some underscores before and after it
        new_cols[col] = '--' + col + '--'
    # rename the column headers
    summary = summary.rename(columns=new_cols)        
    
    # convert the df to a string format table with floats getting a 2 decimal format
    # b = summary.to_markdown(numalign='left', floatfmt=".2f", tablefmt=tablefmt, stralign='left')
    b = summary.to_html(col_space=100, index=False, justify='center')
    # replace the dataframe in the dictionary dict with the viewer friendly table
    dictionary['Summary'] = b
        
    return dictionary


def output_absent_dict(absent_df):
    # get the states present in the absent dataframe
    states = pd.unique(absent_df['Shop'])
    # create an empty dict to sort each state into
    absent_dict = {}
    # go thru each state
    for state in states:
        # create an empty dict for each state so that it can get the df and recipients later
        absent_dict[state] = {}
        # only get the employees for that state
        states_df = absent_df[absent_df['Shop'] == state]
        # only retain the 'ID', "Name", & "Productive" columns
        states_df = states_df[['ID','Name','Productive']]
        # append the df to the dict
        absent_dict[state]['Absent'] = states_df
    
    return absent_dict




#%% This section is for the emailing functions

def email_sub80_results(date_str, state, state_dict):
    service = get_email_service()
    
    # import shutil
    print('Sending ' + state + ' Sub80Direct to: ', state_dict['Recipients'])
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # the directory to save the csv files as a copy to
    directory = Path.home() / 'documents' / 'High_Direct_Hours_Reports'
    if not os.path.exists(directory):
        os.makedirs(directory) 
    # the start of the file name
    file_start = state + ' High Percentage Indirect Hours ' 
    # the ending of the file 
    file_end = ' Report for ' + date_str.replace('/','-')  + '.csv'

    # convert the dataframes to a HTML table for email digestable format
    string_formatted_dict_of_dfs = return_pretty_string_format_of_df_dicts(copy.deepcopy(state_dict))
    # get the summary html df
    summary_html = string_formatted_dict_of_dfs['Summary']
    # get the details html df
    details_html = string_formatted_dict_of_dfs['Detail']
    # Create the first line of the email
    email_start = "\n<p>Summary of Employees with less than 90% Direct Hours</p>\n"
    # create the line break and "Breakdown" 
    email_middle = "\n<br></br>\n<p>Breakdown</p>\n"
    # little note regarding who to contact about the reports
    email_pre_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # create the hyperlink ending that links to a directory
    email_end = '<a href=' + str(directory) +'>Check Here for backups or missed email results</a>'
    # combine all the HTML strings to form the message
    email_msg = email_start + summary_html + email_middle + details_html + email_pre_end + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = state + ' High Percentage Indirect Hours Report: ' + date_str
    # create the sender
    msg['From'] = state + ' High Indirect Percentage Report <' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(state_dict['Recipients'])
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    # save the summary as csv
    file_mid = 'Summary'
    # create the summary_file name
    summary_file = file_start + file_mid + file_end
    
    summary_file_out = directory / summary_file
    # send it to a csv at the destination
    state_dict['Summary'].to_csv(summary_file_out, index=False)
    ## copy the file from the local drive to somewhere on the server
    # shutil.copy(summary_file, "I://Scanned Docs//")
    
    # save the details as a csv
    file_mid = 'Details'
    # create the details_file name
    details_file = file_start + file_mid + file_end
    details_file_out = directory / details_file
    # send it to a csv at the destination
    state_dict['Detail'].to_csv(details_file_out, index=False)
    ## copy the file from the local drive to somewhere on the server
    # shutil.copy(details_file, "I://Scanned Docs//")
    # for each file, add the attachment
    for filename in [summary_file_out, details_file_out]:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(filename, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(filename)))
        msg.attach(part)
    
    # msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')       
    

        


def email_sub2lots_results(date_str, state, state_dict):
    service = get_email_service()
    # import shutil
    print('Sending ' + state + ' Sub2Lots email to: ', state_dict['Recipients'])
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # the directory to save the csv files as a copy to
    directory = Path.home() / 'documents' /'Sub_2_Lots_Reports'
    if not os.path.exists(directory):
        os.makedirs(directory) 
    # the start of the file name
    file_start = state + ' Sub 2 Lots Report ' 
    # the ending of the file 
    file_end = ' Report for ' + date_str.replace('/','-')  + '.csv'

    # convert the dataframes to a HTML table for email digestable format
    string_formatted_dict_of_dfs = return_pretty_string_format_of_df_dicts(copy.deepcopy(state_dict))
    # get the summary html df
    summary_html = string_formatted_dict_of_dfs['Summary']
    # get the details html df
    details_html = string_formatted_dict_of_dfs['Detail']
    # Create the first line of the email
    email_start = "\n<p>Summary of Employees with Less Than 2 Lots Clocked</p>\n"
    # create the line break and "Breakdown" 
    email_middle = "\n<br></br>\n<p>Breakdown</p>\n"
    # little note regarding who to contact about the reports
    email_pre_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # create the hyperlink ending that links to a directory
    email_end = '<a href=' + str(directory) +'>Check Here for backups or missed email results</a>'
    # combine all the HTML strings to form the message
    email_msg = email_start + summary_html + email_middle + details_html + email_pre_end +email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = state + ' Less Than 2 Lots Report: ' + date_str
    # create the sender
    msg['From'] = state + ' Less Than 2 Lots Report <' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(state_dict['Recipients'])
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    # save the summary as csv
    file_mid = 'Summary'
    # create the summary_file name
    summary_file = file_start + file_mid + file_end
    summary_file_out = directory / summary_file
    # send it to a csv at the destination
    state_dict['Summary'].to_csv(summary_file_out, index=False)
    ## copy the file from the local drive to somewhere on the server
    # shutil.copy(summary_file, "I://Scanned Docs//")
    
    # save the details as a csv
    file_mid = 'Details'
    # create the details_file name
    details_file = file_start + file_mid + file_end
    details_file_out = directory / details_file
    # send it to a csv at the destination
    state_dict['Detail'].to_csv(details_file_out, index=False)
    ## copy the file from the local drive to somewhere on the server
    # shutil.copy(details_file, "I://Scanned Docs//")
    # for each file, add the attachment
    for filename in [summary_file, details_file]:
        filepath = directory / filename
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(filepath, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(filename)))
        msg.attach(part)

    # msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')       
    
    

def email_absent_list(date_str, state, state_dict):
    service = get_email_service()
    # import shutil
    print('Sending ' + state + ' Absent email to: ', state_dict['Recipients'])
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # the directory to save the csv files as a copy to
    directory = Path.home() / 'documents' / 'Absent_Reports'
    if not os.path.exists(directory):
        os.makedirs(directory) 
    # the file name
    file_name =  state + ' Absent Report for ' + date_str.replace('/','-')  + '.csv'
    out_file = directory / file_name
    # send the csv file to my local computer for safe keeping
    state_dict['Absent'].to_csv(out_file, index=False)
    # shutil.copy(summary_file, "I://Scanned Docs//")
    
  
    
    # transform the dataframe to a html table for email viewing
    absent_html = state_dict['Absent'].to_html(col_space=100, index=False, justify='center')
    # Create the first line of the email
    email_start = "\n<p>Employees not clocked in or had missed clocks on " + date_str + "</p>\n"
    # little note regarding who to contact about the reports
    email_pre_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"
    # create the hyperlink ending that links to a directory
    email_end = '<a href=' + str(directory) +'>This will be a link to the Z drive with backups of the data at some point</a>'
    # combine all the HTML strings to form the message
    email_msg = email_start + absent_html + email_pre_end + email_end
    # create new message
    msg = MIMEMultipart("html")

    # # create the sender
    msg['From'] = state + ' Absent/Missing Clocks Report <' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(state_dict['Recipients'])
    # create the subject line
    msg['Subject'] = state + ' Absent/Missing Clocks Report: ' + date_str    
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
   
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(out_file, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(file_name)))
    msg.attach(part)
    
    
    # msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message,
        'To': ', '.join(state_dict['Recipients'])
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')       
    
   
def make_text_barchart(df, groupby_list, target_col, char_worth=2, barchart_char='*', how='cascading'):
    ''' 
    how = ['cascading','cumulative','default']
    'cascading': cumulative but the previous data points are filled with underscores instead of the barchar_char
        or it could be said that this is the same as default but with the 'bars' offset by the % already taken up by previous rows
    '''
    x = df.groupby(groupby_list).sum().reset_index()
    
    x['% Share'] = x[target_col] / x[target_col].sum()

    x['% Share big'] = (x['% Share'] * 100).round(0)
       
    # max_chars = int(100 / char_worth)
    
    x['% Share modulo'] = x['% Share big'] % char_worth
    
    x['number chars'] = (x['% Share big'] - x['% Share modulo']) / char_worth
    
    max_chars = int(x['number chars'].sum())
    
    x = x.sort_values(by='Hours', ascending=False).reset_index(drop=True)
    
    x['cumsum number chars'] = x['number chars'].cumsum()
    
    
    x['bar chart'] = [''] * len(x.index)
    if how == 'cascading':
        for i in x.index:
            num_chars = int(x.loc[i,'number chars'])
            if i == 0:
                
                x.loc[i, 'bar chart'] =  num_chars * barchart_char + (max_chars-num_chars) * '_'
            else:
                prefix_filler = int(x.loc[i-1, 'cumsum number chars']) * '_'
                cumulative_chars = int(x.loc[i, 'cumsum number chars'])
                x.loc[i, 'bar chart'] = prefix_filler + num_chars * barchart_char + (max_chars - cumulative_chars) * '_'
    
    elif how == 'cumulative':
        for i in x.index:
            num_chars = int(x.loc[i,'cumsum number chars'])
            x.loc[i,'bar chart'] = num_chars * barchart_char + (max_chars-num_chars) * '_'
            
        
    elif how == 'default':
        for i in x.index:
            num_chars = int(x.loc[i,'number chars'])
            
            x.loc[i, 'bar chart'] =  num_chars * barchart_char + (max_chars-num_chars) * '_'
        
        
    
    
    # get the columns that are needed to be retured
    x = x[groupby_list + [target_col] + ['% Share','bar chart']]    
    # tell the audience how much each character is worth percentage wise
    x = x.rename(columns = {'bar chart': '% Share Bar Chart (' + barchart_char + '=' + str(char_worth) + '%)'})
    # round th e
    x['% Share'] = (x['% Share'].round(2) * 100).astype(int).astype(str) + ' %'
    
    return x





def email_mdi(date_str, state, state_dict, email_dict):
    service = get_email_service()
    recipient_list = email_dict[state] #+ ['rrichard@crystalsteel.net', 'emohamed@crystalsteel.net']
    print('Sending ' + state + ' MDI email to: ', recipient_list)
    br = "<br></br>\n"
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # the directory to save the csv files as a copy to
    directory = Path.home() / 'documents' / 'MDI' / 'Automatic'
    # if the directory does not exist
    if not os.path.exists(directory):
        # create the directory
        os.makedirs(directory) 
    
    # the file name
    file_name = directory / (state + ' MDI ' + date_str.replace('/','-')  + '.xlsx')
    # send the csv file to my local computer for safe keeping
    
   

    # get the df from the dict
    mdi_summary = state_dict['MDI Summary']
    # round off the numbers 
    mdi_summary = mdi_summary.round(2)
    # change the efficiency numbers to strings in percentage format
    # mdi_summary.loc[['Efficiency (Model)', 'Efficiency (Old)']] = (mdi_summary.loc[['Efficiency (Model)', 'Efficiency (Old)']] * 100).round(1).astype(str) + ' %'
    # Ensure the column can hold string values before assignment
    mdi_summary.loc[['Efficiency (Model)', 'Efficiency (Old)']] = (
        (mdi_summary.loc[['Efficiency (Model)', 'Efficiency (Old)']] * 100)
        .round(1)
        .astype(str) + ' %'
    ).astype(object)  # Explicitly convert to object dtype before assignment

    # put prettier names on the series for the email 
    mdi_summary = mdi_summary.rename({'Earned (Model)': 'Earned Hours (EVA)',
                                      'Earned (Old)': 'Earned Hours (HPT)',
                                      'Direct':'Direct Hours Worked',
                                      'Efficiency (Model)':'Efficiency (EVA)',
                                      'Efficiency (Old)': 'Efficiency (HPT)',
                                      'Indirect':'Indirect Hours Worked'})
    
    mdi_summary_html = mdi_summary.to_html(col_space=100, justify='center', header=False)
    
    
    # get the df from the dict
    direct_by_dept = state_dict['Direct by Department'].reset_index()
    
    # get the little barchart of percentage share of hours
    by_job = make_text_barchart(direct_by_dept, ['Job #'], 'Hours', how='cascading')
    by_lot = make_text_barchart(direct_by_dept, ['Job #','Cost Code'], 'Hours', how='cascading')
    by_dept_lot = make_text_barchart(direct_by_dept, ['Job #','Cost Code','Department'], 'Hours', how='cascading')
    by_dept = make_text_barchart(direct_by_dept, ['Department'], 'Hours', how='cascading')
    
    # convert the job # to a nice string for formatting purposes
    by_job['Job #'] = by_job['Job #'].astype(int).astype(str)
    by_lot['Job #'] = by_lot['Job #'].astype(int).astype(str)
    by_dept_lot['Job #'] = by_dept_lot['Job #'].astype(int).astype(str)
    
    
    # send to html format
    by_job_html = by_job.to_html(col_space=100, justify='center', index=False)
    by_lot_html = by_lot.to_html(col_space=100, justify='center', index=False)
    by_dept_lot_html = by_dept_lot.to_html(col_space=100, justify='center', index=False)
    by_dept_html = by_dept.to_html(col_space=100, justify='center', index=False)
    
    
    # Create the first line of the email
    email_start = "\n<p>MDI Data & Daily Recap " + date_str + "</p>\n"
    # little note regarding who to contact about the reports
    email_pre_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"
    # create the hyperlink ending that links to a directory
    email_end = '<a href=' + str(directory) +'>This will be a link to the Z drive with backups of the data at some point</a>'
    # combine all the HTML strings to form the message
    email_msg = email_start + '<u>MDI Summary\n</u>' + mdi_summary_html + br
    email_msg += '<b>See attached Excel file for detailed breakdown of these numbers!</b>' + br + br
    email_msg += '<u>Direct Hours Breakdown by Job\n</u>' + by_job_html + br
    email_msg += '<u>Direct Hours Breakdown by Lot\n</u>' + by_lot_html + br
    email_msg += '<u>Direct Hours Breakdown by Department\n</u>' + by_dept_html + br    
    email_msg += '<u>Direct Hours Breakdown by Lot & Department\n</u>' + by_dept_lot_html + br 
    email_msg += email_pre_end + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = state + ' MDI & Daily Recap Report: ' + date_str
    # create the sender
    msg['From'] = state + ' MDI & Daily Recap Report <' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(recipient_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
   
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_name, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(file_name)))
    msg.attach(part)
    
    
    # msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')    
    
    
    
         


def email_eva_vs_hpt(date_str, eva_vs_hpt_dict, email_recipients):
    service = get_email_service()
    br = "<br></br>\n"
    
    # email_recipients = ['cwilson@crystalsteel.net']
    print('Sending EVA vs HPT email to: ', email_recipients)
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # the directory to save the csv files as a copy to
    # the file name
    file_names = []
    for key in eva_vs_hpt_dict.keys():
        if eva_vs_hpt_dict[key] is None:
            continue
        file_names.append(eva_vs_hpt_dict[key]['Filename'])
    
    # add rounding to some columns
    rounding_dict = {'Tons':2, 'EVA':1, 'HPT':1, 'Hr. Diff':1, '% Diff':2}    
    renaming_dict = {'Tons':'--Tons--', 'EVA':'--EVA--','HPT':'--HPT--', 'Quantity':'--Qty--'}
    
    # init the email_msg
    email_msg = "\n<p>EVA & HPT recap " + date_str + "</p>\n"
    
    # This is because sometimes we dont find anything in fablisting for yesterday
    # but when we do, then do all this stuff
    if not eva_vs_hpt_dict['Yesterday'] is None:
        missing_pieces = eva_vs_hpt_dict['Yesterday']['Missing']
        
        missing_pieces['Job #'] = missing_pieces['Job #'].astype(int).astype(str)
        
        missing_pieces_html = missing_pieces.to_html(col_space=100, index=False, justify='center') 
    
       # get the different dataframes
        by_pcmark = eva_vs_hpt_dict['Yesterday']['Pcmark'].iloc[:25]
        
        percent_diff = by_pcmark['% Diff'].copy()
        by_pcmark = by_pcmark.copy()
        by_pcmark.loc[:,'bins'] = pd.cut(percent_diff, [0,0.5,1,max(2, max(percent_diff))])
        by_pcmark_summary = by_pcmark.groupby('bins', observed=False).count()
        by_pcmark = by_pcmark.drop(columns=['bins'])
        by_pcmark_summary['Range'] = ['< 50%','50-100%','> 100%']
        by_pcmark_summary = by_pcmark_summary.set_index('Range')
        by_pcmark_summary['% Share'] = by_pcmark_summary['Quantity'] / by_pcmark_summary['Quantity'].sum()
        by_pcmark_summary['% Share'] = (by_pcmark_summary['% Share'].round(2) * 100).astype(int).astype(str) + ' %'
        by_pcmark_summary = by_pcmark_summary[['Quantity','% Share']].reset_index()
    
        by_pcmark = by_pcmark.round(rounding_dict)
        by_pcmark_summary_html = by_pcmark_summary.to_html(col_space=100, index=False, justify='center')
    
        by_pcmark = by_pcmark.rename(columns = renaming_dict)
        by_pcmark['% Diff'] = (by_pcmark['% Diff'] * 100).round(0).astype(int).astype(str) + ' %'

        by_pcmark['Job #'] = by_pcmark['Job #'].astype(int).astype(str)
        by_pcmark = by_pcmark.to_html(col_space=100, index=False)

        email_msg += '<u>Pieces missing from the Model files\n</u>' + missing_pieces_html
        
        # combine all the HTML strings to form the message
        email_msg += '<p>Remember to check Fablisting for spelling errors (check Job, Lot, & Pcmark), as that can cause pieces to show up in this report.' + br
        email_msg += "If there are no spelling errors & pieces still show up here, it is likely that there is not a corresponding LOT file in the Dropbox</p>" + br
        email_msg += "Here is the breakdown of how many of yesterday's pieces differ between models" + by_pcmark_summary_html + br
        email_msg += "\n<p>See the attached Excel files for a full breakdown of EVA hours vs. HPT Hours." + br
        email_msg += "The tables in this email only show the top offenders...</p>\n" 
        email_msg += '<u>\nDifference between EVA & HPT models Yesterday, by Piecemark (TOP 25)\n</u>' + by_pcmark + br

    # this is when there isnt anything from fablisting for yesterday
    else:
        email_msg += '<p>Remember to check Fablisting for spelling errors (check Job, Lot, & Pcmark), as that can cause pieces to show up in this report.' + br
        email_msg += f'There were no pieces found in fablisting for {date_str}.</p>' + br
        email_msg += "\n<p>See the attached Excel files for a full breakdown of EVA hours vs. HPT Hours." + br
        email_msg += "The tables in this email only show the top offenders...</p>\n" 
    
    ''' There is an almost gurantee that the following will always have records in them'''     
    
    by_lot = eva_vs_hpt_dict['10 day']['Lot'].iloc[:15]
    by_job = eva_vs_hpt_dict['60 day']['Job'].iloc[:15]
    
    # by_pcmark = by_pcmark.drop(columns=['bins'])
    by_lot = by_lot.round(rounding_dict)
    by_job = by_job.round(rounding_dict)
    # rename some columns
    by_lot = by_lot.rename(columns = renaming_dict)
    by_job = by_job.rename(columns = renaming_dict)
    # convert the percentage column to an actual percentage
    by_lot['% Diff'] = (by_lot['% Diff'] * 100).round(0).astype(int).astype(str) + ' %'
    by_job['% Diff'] = (by_job['% Diff'] * 100).round(0).astype(int).astype(str) + ' %'
    # convert the job # column to int
    by_lot['Job #'] = by_lot['Job #'].astype(int).astype(str)
    by_job['Job #'] = by_job['Job #'].astype(int).astype(str)
    
    
    # covnert to html tables
    by_lot = by_lot.to_html(col_space=100, index=False, justify='center')
    by_job = by_job.to_html(col_space=100, index=False, justify='center')
    

    

    email_msg += '<u>\nDifference between EVA & HPT models Last 10 days, by Lot (TOP 15)\n</u>' + by_lot + br
    email_msg += '<u>\nDifference between EVA & HPT models Last 60 days, by Job (TOP 15)\n</u>' + by_job + br
    email_msg += "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"

    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = 'EVA vs HPT Recap Report: ' + date_str
    # create the sender
    msg['From'] = 'EVA vs HPT Recap Report <' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(email_recipients)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    for key in eva_vs_hpt_dict.keys():
        if eva_vs_hpt_dict[key] is None:
            continue
        file_name = eva_vs_hpt_dict[key]['Filename']
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(file_name, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(file_name)))
        msg.attach(part)
    
    # msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')     
    
    
    
    
def emaIL_attendance_hours_report(date_str, state, file_name, email_dict):
    service = get_email_service()
    
    recipient_list = email_dict[state]# + ['rrichard@crystalsteel.net', 'emohamed@crystalsteel.net']
    
    
    # recipient_list = ['cwilson@crystalsteel.net']
    
    print('Sending ' + state + ' Weekly Attendance Hours Report to: ', recipient_list)
    br = "<br></br>\n"
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    
    email_msg = "\n<p>Weekly Attendance Hours Report for week of " + date_str + "</p>\n"
    email_msg += '<p>Use the "View" tab in the attached Excel file </p>' + br
    email_msg += '<p>Color Coding: </p>'
    email_msg += '<p>Green      = Greater than 48 hours worked</p>'
    email_msg += '<p>Yellow     = 48 hours to 40 hours worked</p>'
    email_msg += '<p>Red        = Less than 40 hours worked</p>'
    email_msg += '<p>Black      = Did not work this week</p>' + br
    email_msg += '<p> The columns are ordered by hours worked, in descending order for the most recent week</p>'
    email_msg += '<p> The employee with the most hours worked in the most recent week will be on the left </p>'
    email_msg += "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"
    
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = 'Weekly Attendance Hours Report ' + state +': ' + date_str
    # create the sender
    msg['From'] = 'Weekly Attendance Hours Report <' + my_email +'>'
    
    if isinstance(recipient_list, str):
        raise ValueError(' The recipient_list variable is a string & needs to be a list ')
    
    # create the receiver
    msg['To'] = ', '.join(recipient_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
   
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_name, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(file_name)))
    msg.attach(part)
    
    
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')           
    
    
def email_delivery_calendar_changelog(date_str, file_name, recipient_list):
    service = get_email_service()
    
    # import shutil
    print('Sending ' + date_str + ' Delivery Calendar change information')
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = 'Delivery Calendar Change data'
    # create the sender
    msg['From'] = 'Delivery Calendar <' + my_email +'>'
    
    if isinstance(recipient_list, str):
        raise ValueError(' The recipient_list variable is a string & needs to be a list ')
    
    email_msg = "\n<p>Delivery Calandar Changelog " + date_str + "</p>\n"
    df = pd.read_csv(file_name)
    df = df.sort_values(by=['action','shop','type_of_work'])
    df_html = df.to_html(col_space=100, index=False)
    email_msg += df_html
    
    # create the receiver
    msg['To'] = ', '.join(recipient_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
   
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_name, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(file_name)))
    msg.attach(part)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')  
    
    

def email_error_message(error_messages_list, recipient_list='cwilson@crystalsteel.net'):
    service = get_email_service()
    
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = 'Emailing Error!'
    # create the sender
    msg['From'] = 'errors <' + my_email +'>'
    
    if isinstance(recipient_list, str):
        recipient_list = [recipient_list]
    
    email_msg = "\n<p>"
    for i in error_messages_list:
        email_msg += '<br>' + i +'<br>'
    email_msg += "</p>\n"
  
    
    # create the receiver
    msg['To'] = ', '.join(recipient_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
   
    part = MIMEBase('application', "octet-stream")
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')  