# -*- coding: utf-8 -*-
"""
Created on Fri Feb 18 13:36:53 2022

@author: CWilson
"""

import sys
import datetime
import pandas as pd
import time
import numpy as np
from Lots_schedule_calendar.cal_setup import get_calendar_service
from utils.attendance_google_sheets_credentials_startup import init_google_sheet
from Lots_schedule_calendar.manila_calendar_emailing_functions import send_new_work_type_notice_email, send_date_change_notice_email
from Lots_schedule_calendar.LOTS_shipping_schedule_conversion2 import get_shipping_schedule


_SendEmails = True



manila_emails = ['cwilson@crystalsteel.net', 'nmarinduque@crystalsteel.net', 
                 'aanonuevo@crystalsteel.com', 'manila_purchasing@crystalsteel.net']
# manila_emails = ['cwilson@crystalsteel.net']


manila_cal_id = 'c_8a519fn21q0j3ot6q3n4g3p1lg@group.calendar.google.com'
cal_id = manila_cal_id



def manila_calendar_as_dict(cal_id):
    try:
        # initialize the calendar service       
        service = get_calendar_service()
        # starting time, only find events after right now
        now = datetime.datetime.now(datetime.UTC).isoformat()
        
        events_result = service.events().list(
               calendarId = cal_id,
               singleEvents = True,
               orderBy = 'startTime',
               maxResults = 2500).execute()
        
        events_dict = events_result.get('items', [])
        
    
    except:
        print('could not get calendar')     
        
    return events_dict
    

def manila_calendar_as_df(cal_id):
    
    
    i = 0
    while 'events_dict' not in locals() and i < 10:
        
        events_dict = manila_calendar_as_dict(cal_id)
        time.sleep(10)
        i += 1
     
    
    
    
    
    if len(events_dict) != 0:
        # convert the dict to a dataframe
        events_df = pd.DataFrame().from_dict(events_dict)
        date_series = events_df['start'].apply(pd.Series)
        to_datetime_col = date_series.columns[0]
        # the 'start' column is a bunch of dict's, so convert it to series
        events_df = events_df.join(date_series)
        # convert the date to datetime
        events_df[to_datetime_col] = pd.to_datetime(events_df[to_datetime_col])
    else:
        events_df = {}
    
    return events_df



def create_event(cal_id, name, date, description='', color_id=8):
    ''' date is in datetime.datetime() format '''
   
    # initialize the calendar service 
    service = get_calendar_service()
    
    event_result = service.events().insert(
        calendarId = cal_id,
        body={
            "summary": name,
            "description": description,
            # "start": {"date": date.isoformat()},
            # "end": {"date": (date+datetime.timedelta(hours=1)).isoformat()},
            "start": {"dateTime": date.isoformat(), 'timeZone': 'Asia/Manila'},
            "end": {"dateTime": (date+datetime.timedelta(hours=1)).isoformat(), 'timeZone': 'Asia/Manila'},            
            "colorId": color_id
        }
    ).execute()    
    
    print("created event:")
    print("\tid: ", event_result['id'])
    print("\tname: ", event_result['summary'])
    print("\ton date: ", event_result['start']['dateTime'], event_result['start']['timeZone'])  



def update_event(cal_id, event_id, summary, date, description, color_id=8):

    # initialize the calendar service           
    service = get_calendar_service()
    
    event_result = service.events().update(
           calendarId = cal_id,
           eventId = event_id,
           body = {
            "summary": summary,
            "description": description,
            "start": {"dateTime": date.isoformat(), 'timeZone': 'Asia/Manila'},
            "end": {"dateTime": (date+datetime.timedelta(hours=1)).isoformat(), 'timeZone': 'Asia/Manila'},  
            # "colorId": color_id
            },
         ).execute()    
    
    print('updated event:')
    print("\tid: ", event_result['id'])
    print("\tname: ", event_result['summary'])
    print("\ton date: ", event_result['start']['dateTime'], event_result['start']['timeZone'])  



def delete_manila_event(cal_id, event_id):
    # initialize the calendar service           
    service = get_calendar_service()
    # delete    
    service.events().delete(calendarId=cal_id, eventId=event_id).execute()

#%%

def manila_calendar_function():
    
    
    # get the current time
    now = datetime.datetime.now()
    
    NUMBER_OF_WEEKS_THAT_ARE_ALLOWABLE = 8
    
    cutoff = now - datetime.timedelta(days=7*NUMBER_OF_WEEKS_THAT_ARE_ALLOWABLE)
    
    ''' Current calendar events'''
    events_df = manila_calendar_as_df(manila_cal_id)
    
    
    try:
        ''' delete events that are older then 8 week before today '''
        to_delete = events_df[events_df['dateTime'].dt.tz_localize(None) < cutoff]
        
        # to_delete = this_seq
        for idx in to_delete.index:
            row = to_delete.loc[idx]
            print(row)
            delete_manila_event(manila_cal_id, event_id = row.squeeze()['id'])
        ''' email out the list of deleted events '''
        deleted_summaries = list(to_delete['summary'])
    except:
        print('could not find anything to delete')
    
    events_df = manila_calendar_as_df(manila_cal_id)
    if len(events_df):
        ''' this is where I will only process events that have MY schema '''
        # get the job from the first 4 digits of the title
        events_df['Job'] = events_df['summary'].str[:4]
        # convert the job column to number
        events_df['Job'] = pd.to_numeric(events_df['Job'], errors='coerce')      
        # get rid of anything where the job is not a number
        events_df = events_df[~events_df['Job'].isna()]
        # get the sequence number from the title
        events_df['Sequence'] = events_df['summary'].str.split(' ', expand=True)[2]
        # convert the sequence number tonumber
        events_df['Sequence'] = pd.to_numeric(events_df['Sequence'])
        # get the reminder type from the title
        events_df['Reminder #'] = events_df['summary'].str.split('- ', expand=True)[1]



    
    # scheduler_column = 'Delivery'
    scheduler_column = 'Dwgs Needed in Shop'
    
    ''' Shipping schedule '''
    # get the current shipping schedule for all shops & only for sequences
    ss = get_shipping_schedule(shop=None, type_of_work='Sequence')
    # rename the scheduler column to 'scheduler'
    ss = ss.rename(columns={scheduler_column:'scheduler'})
    # conver the delivery column to datetime
    ss['scheduler'] = pd.to_datetime(ss['scheduler'], errors='coerce')
    # get rid of any rows that didn't convert correctly
    ss = ss[~ss['scheduler'].isna()]
    # get rid of rows where delivery is over x weeks ago
    ss = ss[ss['scheduler'] > now - datetime.timedelta(days=7*(NUMBER_OF_WEEKS_THAT_ARE_ALLOWABLE-6))]
    # convert the job column to a number
    ss['Job'] = pd.to_numeric(ss['Job'], errors='coerce')
    # get rid of any non compliant job numbers
    ss = ss[~ss['Job'].isna()]
    # get rid of any job numbers that are outside of 1900-9000
    ss =ss[(ss['Job'] > 1900) & (ss['Job'] < 9000)]
    # convert the sequqnce numbers to number
    ss['Number'] = pd.to_numeric(ss['Number'], errors='coerce')
    # get rid of any non compliant sequqnce numbers
    ss = ss[~ss['Number'].isna()]
    # create a column that converts the decimal sequeqnce number to just the integer version
    ss['int Number'] = ss['Number'].astype(int)
    
 
    x1 = ss.groupby(['Job', 'int Number']).min()
    x2 = ss.groupby(['Job', 'int Number']).max()
    x1['scheduler_max'] = x2['scheduler']
    x1 = x1[x1['scheduler'] != x1['scheduler_max']]
    
    # this gets all the different job-sequences and earliest delivery dates with a unique sequence number column
    ss1 = ss.groupby(['Job', 'int Number']).min()['scheduler']
    # put the index to columns
    ss1 = ss1.reset_index()
    # sort by delivery date
    ss1 = ss1.sort_values('scheduler')
    # calculate the first reminder date - 8 weeks
    ss1['1st Reminder'] = ss1['scheduler'] - datetime.timedelta(days = 7 * 2)
    # calculate the second reminder date - 6 weeks
    ss1['2nd Reminder'] = ss1['scheduler'] - datetime.timedelta(days = 7 * 1)
    #calculate the final reminder date - 4 weeks
    ss1['FINAL Reminder'] = ss1['scheduler'] - datetime.timedelta(days = 7 * 0)
    # get rid of rows where 1st reminder is older than the cutoff
    ss1 = ss1[ss1['1st Reminder'] > cutoff]
    
    
    

    reminders = ['1st Reminder','2nd Reminder','FINAL Reminder']
    
    now_str_timestamp = '(' + now.strftime('%Y-%m-%d %H:%M') + ') '
    final_to_delete = pd.DataFrame()
    for idx in ss1.index:
        
        # get just that job & delivery date
        row = ss1.loc[idx]
        # get the job #
        job = row['Job']
        # get the sequence number
        seq = row['int Number']
        
        scheduler_date_str = row['scheduler'].strftime('%Y-%m-%d')
        
        print(job, seq)
        if len(events_df):
            # find the calendar events for this job
            this_seq = events_df[events_df['Job'] == job]
            # find the calendar events for this sequence
            this_seq = this_seq[this_seq['Sequence'] == seq]
            # get the unique list of the reminders to make sure we have all 3 present on the calendar
            these_reminders = list(set(this_seq['Reminder #']))
        else:
            these_reminders = []
        
        
        if (row[reminders[0]] - now).days > 120:
            print('First reminder is further than 120 days in future')
            continue
        
        
        # go thru each of the 3 main types of reminders  via reminders list
        for reminder in reminders:
            # get the shipping schedule date for the 
            ss_date = row[reminder] + datetime.timedelta(hours=10)     
            
            
            # create the email name
            email_name = str(job) + ' Seq. ' + str(seq)
            # create the event name AKA summary
            event_name = email_name + ' - ' + reminder
            # if the reminder is present on the calender, check to see if it needs the date updated
            if reminder in these_reminders:
                # get the event for the job-sequence-reminder 
                cal_event = this_seq[this_seq['Reminder #'] == reminder]
                # error handling when there are multiple calendar events for a sequence-reminder
                if cal_event.shape[0] == 1:
                    cal_event = cal_event.squeeze()
                else:
                    print('THERE ARE MULTIPLE CALENDAR EVENTS FOR: {} Seq. {} - {}'.format(job, seq, reminder))
                    ''' send error notice email out '''
                    final_to_delete = final_to_delete.append(this_seq)
                    # input('FIX THIS JOHHNY.......')
                    
                    for idx in this_seq.index:
                        row = this_seq.loc[idx]
                        print('deleting records for {} Seq. {} - {}'.format(job,seq,reminder))
                        delete_manila_event(manila_cal_id, event_id = row.squeeze()['id'])   
                        
                    print('\t\t\tRECURSION')
                    manila_calendar_function()
                    
                    pass
                
                # get the date from the calendar event
                cal_date = cal_event['dateTime']
                # if the calendar date & the SS date don't match then change the calender event to the SS date
                if cal_date.date() != ss_date.date():
                    # get the event id
                    event_id = cal_event['id']
                    # get the event name
                    event_name = cal_event['summary']
                    # get the html link
                    event_html_link = cal_event['htmlLink']
                    # get the shipping schedule date as a string
                    ss_date_str =  ss_date.strftime('%Y-%m-%d')
                    # turn the calendar date to a string for the description
                    cal_date_str = cal_date.strftime('%Y-%m-%d')
                    # get the calendar description
                    cal_desc = cal_event['description']
                    # change the dwgs needed in shop date
                    cal_desc_split = cal_desc.split('\n')
                    # change the date in the first line of the description
                    cal_desc_split[0] = cal_desc_split[0][:-10] + scheduler_date_str
                    # rejoin the split description into one string seperated by newlines
                    cal_desc_new = '\n'.join(cal_desc_split)
                    # add the new lines to the description
                    updated_desc = cal_desc_new + now_str_timestamp + 'Reminder date change: ' + cal_date_str + ' to ' + ss_date_str +'\n'
                    #update the event to change the date
                    update_event(manila_cal_id, event_id, summary=event_name, date=ss_date, description=updated_desc)
                    time.sleep(1)
                    if reminder == '1st Reminder' and _SendEmails:
                        print('Sending email notice of date change: {} Seq. {} - {}'.format(job, seq,reminder))
                        send_date_change_notice_email(shop = str(job),
                                                      type_of_work = 'Sequence',
                                                      name = email_name,
                                                      date = scheduler_date_str,
                                                      desc = updated_desc,
                                                      event_url_link = event_html_link,
                                                      recipients_list = manila_emails)
                    
                else:
                    print('The calendar event for {} Seq. {} - {} is correctly date'.format(job,seq,reminder))
           
            # if the reminder is NOT present on the calender, create the event
            else:
                
                # create the description for the new event
                description = '"' + scheduler_column + '" date: ' + scheduler_date_str + '\n'
                description += now_str_timestamp + reminder + ' created\n'
                # create the calendar event
                
                
                
                create_event(manila_cal_id, name=event_name, date=ss_date, description=description)
                time.sleep(1)
                
                new_events_df = manila_calendar_as_df(manila_cal_id)
                # get the html link to that lot
                new_events_html_link = new_events_df[new_events_df['summary'] == event_name].squeeze()['htmlLink']                
                # send email out alerting new sequence on SS
                if reminder == '1st Reminder' and _SendEmails:
                    print('Sending email notice of new sequence: {} Seq. {} - {}'.format(job, seq, reminder))
                    send_new_work_type_notice_email(shop = str(job),
                                                    type_of_work = 'Sequence', 
                                                    name = email_name, 
                                                    date = scheduler_date_str, 
                                                    desc = description,
                                                    event_url_link = new_events_html_link,
                                                    recipients_list = manila_emails)
                    
