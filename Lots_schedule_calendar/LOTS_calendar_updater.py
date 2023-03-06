# -*- coding: utf-8 -*-
"""
Created on Wed Jun 16 10:31:11 2021

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\users\\cwilson\\documents\\python\\Lots_schedule_calendar')
import datetime
import pandas as pd
import time
import numpy as np
import os
from cal_setup import get_calendar_service
from LOTS_shipping_schedule_conversion2 import draw_the_rest_of_the_horse
from LOTS_shipping_schedule_conversion2 import retrieve_from_prod_schedule
from calendar_error_producer_function import produce_error_file
from calendar_change_producer_function import produce_change_file
from calendar_emailing_functions_with_gmail_api import send_new_work_type_notice_email
from calendar_emailing_functions_with_gmail_api import send_date_change_notice_email
from calendar_emailing_functions_with_gmail_api import send_seq_change_notice_email
from calendar_emailing_functions_with_gmail_api import send_error_notice_email
from calendar_emailing_functions_with_gmail_api import daily_changes_new_work_email
from calendar_emailing_functions_with_gmail_api import daily_errors_email
from daily_changes_and_additions_for_whiny_mishler import changes_and_new_work
from daily_changes_and_additions_for_whiny_mishler import daily_errors_summary

from manila_calendar import manila_calendar_function

_SendEmails = False


def get_events_dict(calendar_id):
    
    try:
        # initialize the calendar service       
        service = get_calendar_service()
        # starting time, only find events after right now
        now = datetime.datetime.utcnow().isoformat() + 'Z'
        
        events_result = service.events().list(
               calendarId = calendar_id,
               singleEvents = True,
               orderBy = 'startTime',
               maxResults = 2500).execute()
        
        events = events_result.get('items', [])
        
        # for event in events:
        #     start = event['start'].get('dateTime', event['start'].get('date'))
        #     print(start, event['summary']) 
    
        return events
    
    except Exception as e:
        error_name = 'Calendar List Events Error'
        
        produce_error_file(exception_as_e = e,
                           shop = shop,
                           file_prefix = error_name)      
        
        send_error_notice_email(shop, error_name, e)


def convert_events_dict_to_df(cal_id, type_of_work):
    events_dict = get_events_dict(cal_id)
    if len(events_dict) != 0:
        # convert the dict to a dataframe
        events_df = pd.DataFrame().from_dict(events_dict)
        # only keep the events that have the word 'onsite' in them
        events_df = events_df[events_df['summary'].str.contains(type_of_work)]
        # grab the lots name by dropping the 'Onsite'
        events_df[type_of_work] = events_df['summary'].str[len(type_of_work)+1:]
        # the 'start' column is a bunch of dict's, so convert it to series
        events_df['date'] = events_df['start'].apply(pd.Series)
        # convert the date to datetime
        events_df['date'] = pd.to_datetime(events_df['date'])
    else:
        events_df = pd.DataFrame(columns = ['kind', 'etag', 'id', 'status', 'htmlLink', 'created', 'updated',
                                           'summary', 'description', 'creator', 'organizer', 'start', 'end',
                                           'iCalUID', 'sequence', 'reminders', 'eventType',
                                           'colorId', 'date'])
        events_df[type_of_work] = pd.Series(name=type_of_work, dtype=str)
    
    return events_df




def create_event(calendar_id, name, date, description='', color_id=8):
    ''' date is in datetime.date() format '''
    
    # initialize the calendar service 
    service = get_calendar_service()
    
    event_result = service.events().insert(
        calendarId = calendar_id,
        body={
            "summary": name,
            "description": description,
            "start": {"date": date.isoformat()},
            "end": {"date": (date+datetime.timedelta(days=1)).isoformat()},
            "colorId": color_id
        }
    ).execute()    
    
    print("created event:")
    print("\tid: ", event_result['id'])
    print("\tname: ", event_result['summary'])
    print("\ton date: ", event_result['start']['date'])  
    
    
    
def update_event(calendar_id, event_id, summary, date, description, color_id):

    # initialize the calendar service           
    service = get_calendar_service()
    
    event_result = service.events().update(
           calendarId = calendar_id,
           eventId = event_id,
           body = {
            "summary": summary,
            "description": description,
            "start": {"date": date.isoformat()},
            "end": {"date": (date+datetime.timedelta(days=1)).isoformat()},
            "colorId": color_id
            },
         ).execute()    
    
    print('updated event:')
    print("\tid: ", event_result['id'])
    print("\tname: ", event_result['summary'])
    print("\ton date: ", event_result['start']['date'])  
    
     
def delete_singular_event(calendar_id, event_id):
    # initialize the calendar service           
    service = get_calendar_service()
    # delete    
    service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        
    


def delete_all(calendar_id, event_ids):
    
    '''
    try getting the event_ids with this: 
        calendar_id = shop_dict['CSM']['cal_id']
        calendar_id = shop_dict['CSF']['cal_id']
        calendar_id = shop_dict['FED']['cal_id']
        
        # all events
        event_ids = [i['id'] for i in get_events_dict(calendar_id)]
        # lots only
        event_ids = [i['id'] for i in get_events_dict(calendar_id) if 'LOT' in i['summary']]
        # tickets only
        event_ids = [i['id'] for i in get_events_dict(calendar_id) if 'Ticket' in i['summary']]
        # items only
        event_ids = [i['id'] for i in get_events_dict(calendar_id) if 'Item' in i['summary']]
        # buyouts only
        event_ids = [i['id'] for i in get_events_dict(calendar_id) if 'Buyout' in i['summary']]
        # all misc - no lots
        event_ids = [i['id'] for i in get_events_dict(calendar_id) if 'Ticket' in i['summary'] or 'Buyout' in i['summary'] or 'Item' in i['summary']]
    
    
    event_ids = event_ids[::-1]
    
    '''
    count = len(event_ids)
    print('There are {} calendar events being deleted'.format(count))
    proceed = input(print('Reply (Y) to continue with deleting'))
    if proceed == 'Y':
        i = 0
        success = 0
        failure = 0
        service = get_calendar_service()
        for idnum in event_ids:
            i += 1
            time.sleep(1)
            
            try:
                service.events().delete(calendarId=calendar_id, eventId=idnum).execute()
                success += 1
            except:
                print('failed to delete: {}'.format(idnum))   
                failure += 1
            
            print('{} of {} with {} successes'.format(i,count,success))
                

def get_color(calendar_slice):
    try:
        color = int(calendar_slice['colorId'])
    except:
        color = 8 # return grey if failure
    return color


#This loop mass creates all the lots in the lots_df 
#it is only needed when you delete all of the events for a shop!!!!!   
#You will have to change the job/color thing
# for lot in lots_df['LOT']:
#     squeeze = lots_df[lots_df['LOT'] == lot].squeeze()
#     date = squeeze['Earliest Delivery Date'].date()
#     job = lot[:4]
#     if job == '2003':
#         color = 5
#     elif job == '2009':
#         color = 7
#     else: 
#         color = 2
#     desc = 'Sequences: ' + squeeze['Sequences']
#     if not pd.isna(squeeze['Comment']):
#         desc = desc + '\n' + squeeze['Comment']
#     create_event(cal_id, lot+' Onsite', date, description=desc, color_id=color)
#     time.sleep(1)


def add_to_change_log_v2(shop = '', job = '', type_of_work = '', number = '', action = '', description=''):
    change_log_dir = 'C:\\Users\\cwilson\\Documents\\Python\\Lots_schedule_calendar\\Change_Logs_v2\\'
    now = datetime.datetime.now()
    now_str = now.strftime('%Y-%m-%d %H:%M')
    today = now.date()
    today_str = today.strftime('%Y-%m-%d')
    today_file = change_log_dir + today_str + '.csv'
    headers =  'datetime,shop,job,type_of_work,number,action,description'
    
    joining_list = [now_str, shop, job, type_of_work, number, action, description]
    joining_list_processed = []
    for i in joining_list:
        if i == '':
            joining_list_processed.append('-')
        elif isinstance(i, str):
            joining_list_processed.append(i)
        else:
            try:
                joining_list_processed.append(str(i))
            except:
                joining_list_processed.append('could not convert to string')
            
    if len(joining_list_processed) != len(headers.split(',')):
        print('ERROR joingin list is not the same length as the number of headers')
    
    line_to_file = ','.join(joining_list_processed)
    
    if os.path.exists(today_file):
        with open(today_file, 'a') as f:
            f.write(line_to_file + '\n')
    else:
        with open(today_file, 'x') as f:
            f.write(headers + '\n')
            f.write(line_to_file + '\n')
    
    f.close()

def get_lots_information(shop):

    # get the list of lots and their dates created as an excel file
    draw_the_rest_of_the_horse(shop, send_emails = _SendEmails)
    
    ''' put in a try-catch to get the lots_df file with todays date and then 
    if that fails then just get the most recent file for that shop 
    but still write an error file to keep record of it happening'''
    
    # get today's date to know which lot information file to pull
    today = datetime.datetime.now().strftime("%m-%d-%Y")    
    # cerate filepath
    file_path = 'c://users/cwilson/documents/LOT_schedule_dump/todays_lot_info_' + today +' ' + shop + '.xlsx'
    # process the excel file into a df
    lots_df = pd.read_excel(file_path)
    # rename the index to the lot
    lots_df = lots_df.rename(columns={'index':'LOT'})
    # convert the date to a datetime
    lots_df['Earliest Delivery'] = pd.to_datetime(lots_df['Earliest Delivery'].copy())
    
    return lots_df


#%% Define the dict for each shop


shop_dict = {'CSM':
                 {'cal_id':'c_uqopq4705q7o473uveoah08h7o@group.calendar.google.com',
                  'Recipients':['awhitacre@crystalsteel.net',
                               'bwelsandt@crystalsteel.net',
                               'dwalden@crystalsteel.net',
                               'jgromadzki@crystalsteel.net',
                               'cwilson@crystalsteel.net',
                               'jturner@crystalsteel.net',
                               'rhagins@crystalsteel.net']},
            'FED':
                 {'cal_id':'c_pd5da50k6ci6vsrh8g6a5fvjig@group.calendar.google.com',
                  'Recipients':['cwilson@crystalsteel.net',
                                'lduchesne@crystalsteel.net',
                                'jturner@crystalsteel.net']},
            'CSF':
                 {'cal_id':'c_mrq30egkjgvq6risq4u1epf8e0@group.calendar.google.com' ,
                  'Recipients':['cwilson@crystalsteel.net',
                                'vtalladivedula@crystalsteel.net',
                                'jmixon@crystalsteel.net',
                                'jturner@crystalsteel.net']}}




''' DEBUG SCRIPT 

shop_dict['CSM']['Recipients'] = ['cwilson@crystalsteel.net']
shop_dict['FED']['Recipients'] = ['cwilson@crystalsteel.net']
shop_dict['CSF']['Recipients'] = ['cwilson@crystalsteel.net']


'''

all_recipients = list(set(shop_dict['CSM']['Recipients'] + 
                          shop_dict['FED']['Recipients'] + 
                          shop_dict['CSF']['Recipients'])) + ['mmishler@crystalsteel.com']

#%% This runs the calendar for the misc work types



misc_types = ['Ticket','Item','Buyout']
shop = 'CSM'

for shop in shop_dict.keys():
    print(shop)
    for type_of_work in misc_types:
        # define the calendar id for the shop based on the shop_dict
        cal_id = shop_dict[shop]['cal_id']
        # define the list of recipients from the shop dict for that shop
        recipients_list = shop_dict[shop]['Recipients']    
        # get all of the production schedule items that have work_type = type_of_work
        ss_all = retrieve_from_prod_schedule(shop, type_of_work, send_emails = _SendEmails)
        if ss_all is None:
            print('The variable ss_all is NoneType')
            break
        # create a column for 'JOB#-WORKTYPE#'
        ss_all[type_of_work] = ss_all['Job'].astype(int).astype(str) + '-' + ss_all['Number']
        # get only the open items from ss_all - those not shipped yet
        ss_open = ss_all[ss_all['Shipped?'] == '']
        
        try:
            # get the current calendar events for that work_type
            events_df = convert_events_dict_to_df(cal_id, type_of_work)

                    
            now = datetime.datetime.now()
            ''' delete events that are older then 8 week before today '''
            to_delete = events_df[events_df['date'] < now - datetime.timedelta(days=7*8)]
            # if there is no description in the calendar even tthen we cannot use it and it will cause an error so delete it
            to_delete = to_delete.append( events_df[events_df['description'].isna()] )
            # group by item/ticket/buyout name & count number of entries
            duplicates_to_delete = events_df[[type_of_work, 'id']].groupby([type_of_work]).count()
            # get the ones with > 1 entry
            duplicates_to_delete = duplicates_to_delete[duplicates_to_delete['id'] > 1]
            # add those duplicates to the list to be deleted
            to_delete = to_delete.append(events_df[events_df[type_of_work].isin(list(duplicates_to_delete.index))])
            
            if to_delete.shape[0] == 0:
                print('No {}s need deleting from the calendar'.format(type_of_work))
            else:
                print('There are {} {}(s) events to be deleted from the calendar'.format(to_delete.shape[0],type_of_work))
            
            for idx in to_delete.index:
                row = to_delete.loc[idx].squeeze()
                print(row)
                # i dont knwo why this line was in here - might have checked if there were duplicates of an event?
                # if row.shape[0] != 1:
                #     print('COULD NOT DELETE b/c TOO MANY ROWS: ' + row['summary'][0])
                #     continue
                delete_singular_event(cal_id, row['id'])
                # pause
                time.sleep(1)
            
            # grab the events_df again after deleting the stuff
            events_df = convert_events_dict_to_df(cal_id, type_of_work)
            # get current timestamp for calendar update timestamping
            now_timestamp = '(' + datetime.datetime.now().strftime('%m/%d/%Y %H:%M') + ')'  
            # first find the work_type that are in the ss_open but not in the calendar
            missing_work = np.setdiff1d(ss_open[type_of_work], events_df[type_of_work])
            if missing_work.shape[0] == 0:
                print('No {}s missing from the calendar after considerations'.format(type_of_work))
            else:
                print('There are {} missing {}(s) from the calendar after considerations'.format(missing_work.shape[0],type_of_work))
            
            for work in missing_work:
                cal_name = type_of_work + ' ' + work
                # get the data from the 
                this_work = ss_open[ss_open[type_of_work] == work].squeeze()
                # get the delivery date as a datetime.date()
                date = this_work['Delivery']
                # get the sequences as the description
                desc = '<a href="' + this_work['GoogleSheetLink'] + '">' + 'Link to Google Sheets' + "</a>"
                desc += '<br>' + '(PM: ' + this_work['PM'] + ') ' + this_work['Work Description']
                # include a tag for when the LOT was added to the calendar
                desc += '<br>' + now_timestamp + ' Added to Calendar with delivery date: '+ date.strftime('%m/%d/%Y')
                # create the event
                create_event(cal_id, cal_name, date, desc, color_id=11) # 11 is red as of Nov 14 2022
                # create a string of details for the change file
                change_details = cal_name + '<br>' + date.isoformat() + '<br>' + desc
                # create a file to denote the creation 
                produce_change_file(change_details, shop, work, type_of_work + ' Added')
                add_to_change_log_v2(shop = shop, 
                                     job = work.split('-')[0],
                                     type_of_work = type_of_work,
                                     number = work,
                                     action = 'New Calendar Event',
                                     description = '')
                # wait incase of API limits
                time.sleep(1)
                # geta a new df of the events on the calendar
                new_events_df = convert_events_dict_to_df(cal_id, type_of_work)
                # get the html link to that lot
                new_events_html_link = new_events_df[new_events_df['summary'] == cal_name].squeeze()['htmlLink']
                # send an email out with the new lot information
                if _SendEmails:
                    send_new_work_type_notice_email(shop, type_of_work, work, date.strftime('%m/%d/%Y'), desc, 
                                          new_events_html_link, recipients_list)
                
            ''' Check for discrepancies in the dates '''
            
            cal_dates = pd.DataFrame(data=list(events_df['date']), index=events_df[type_of_work], columns=['cal date'])
            # create a df that is just the lot & dates from the production schedule
            ps_dates = pd.DataFrame(data=list(ss_all['Delivery']), index=ss_all[type_of_work], columns=['ps date'])
            # join the production schedule dates onto the calendar df
            updated_dates = cal_dates.join(ps_dates)
            # in case of duplicates - keep the max date
            # first get the index as a column
            updated_dates = updated_dates.reset_index(drop=False)
            # sort by ps date with max dates on top
            updated_dates = updated_dates.sort_values('ps date', ascending=False)
            #drop duplicate works - keeping the max
            updated_dates = updated_dates.drop_duplicates(type_of_work, keep='first')
            # then reset the index back
            updated_dates = updated_dates.set_index(type_of_work)
            # only keep rows that the calendar date and production schedule date do not align
            updated_dates = updated_dates[updated_dates['cal date'] != updated_dates['ps date']]
            # get rid of any NaT dates from the production schedule -> this happens after the sequences are completed
            updated_dates = updated_dates[~updated_dates['ps date'].isna()]
            
            if updated_dates.shape[0] == 0:
                print('No {}s need the dates updated'.format(type_of_work))
            else:
                print('There are {} {}(s) needing dates changed from the calendar after considerations'.format(updated_dates.shape[0],type_of_work))
            # go thru each lot in the updated_lots df
            for work in pd.unique(updated_dates.index):
                # get the current event information for that lot
                calendar_slice = events_df[events_df[type_of_work] == work].squeeze()
                ''' This if statement should not be needed with the to_delete adding duplicate entry handling'''
                # if there are duplicate entries for something that needs a date updated it will be a dataframe not a series
                if isinstance(calendar_slice, pd.DataFrame):
                    # iterate through each row of the df and delete the records from the calendar
                    for i in calendar_slice.index:
                        row = calendar_slice.loc[i].squeeze()
                        print(row)
                        delete_singular_event(cal_id, row['id'])
                    continue
                # get the current calendar date as iso-format
                cal_date_str = calendar_slice['date'].date().strftime('%m/%d/%Y')
                # get the date from the lots_df AKA the production schedule
                ps_date = ss_all[ss_all[type_of_work] == work].squeeze()['Delivery']
                if not isinstance(ps_date, datetime.date):
                    ps_date = updated_dates[updated_dates.index == work].squeeze()['ps date']
                # conver the production schedule date to a iso-format
                ps_date_str = ps_date.strftime('%m/%d/%Y')
                # the string that gets added to the description so that you can see a history of the change
                append_to_desc = now_timestamp + ' Delivery date changed from ' + cal_date_str + ' to ' + ps_date_str
                # get the description that is currently in the calendar event
                calendar_desc = calendar_slice['description']
                #
                this_work = ss_open[ss_open[type_of_work] == work].squeeze()
                # if there is nothign in the ss_open df then go to ss_all
                if not this_work.shape[0]:
                    this_work = ss_all[ss_all[type_of_work] == work].squeeze()
                    if isinstance(this_work, pd.DataFrame):
                        this_work = this_work[this_work['Delivery'] == ps_date].squeeze()
                    
                desc_first_line = '<a href="' + this_work['GoogleSheetLink'] + '">' + 'Link to Google Sheets' + "</a>"         
                # split the description based on the linebreak
                calendar_desc_list = calendar_desc.split('<br>')
                # change the hyperlink
                calendar_desc_list[0] = desc_first_line
                # join the description back with linebreaks
                calendar_desc = '<br>'.join(calendar_desc_list)                
                # if there is no descirption (unlikely) then just set the append_to_desc as the desc
                if pd.isna(calendar_desc):
                    new_desc = desc_first_line + '<br>' + append_to_desc
                # if there is a description already on the event, then append the append_to_desc
                else:
                    new_desc = calendar_desc + '<br>' + append_to_desc
                
                
                
                # udpate the event with the new date (ps_date) & desc (new_desc), maintain everything else 
                update_event(calendar_id = cal_id, 
                              event_id = calendar_slice['id'], 
                              summary = calendar_slice['summary'],
                              date = ps_date,
                              description = new_desc,
                              color_id = get_color(calendar_slice))        
                
                
                # create a string of details for the change file
                change_details = calendar_slice['id'] + '<br>' + calendar_slice['summary'] + '<br>' + new_desc
                # create a file to denote the change of dates
                produce_change_file(change_details, shop, work, type_of_work + ' Added')
                add_to_change_log_v2(shop = shop, 
                                     job = work.split('-')[0],
                                     type_of_work = type_of_work,
                                     number = work,
                                     action = 'Change Event Date',
                                     description = cal_date_str + ' -> ' + ps_date_str)
                
                if _SendEmails:
                    # send out an email regarding the change in dates
                    send_date_change_notice_email(shop, 
                                                  type_of_work, 
                                                  work, 
                                                  ps_date.strftime('%m/%d/%Y'), 
                                                  new_desc, 
                                                  calendar_slice['htmlLink'], 
                                                  recipients_list)
                
                time.sleep(1)    
            
            
            events_df = convert_events_dict_to_df(cal_id, type_of_work)

            ''' update the hyperlinks '''
            # go through every event of this type on the calendar
            for i in events_df.index:
                # get the calendar data from events_df
                calendar_slice = events_df.loc[i].squeeze()
                # get the information from the shipping schedule
                ss_item = ss_all[ss_all[type_of_work] == calendar_slice[type_of_work]].squeeze()
                # split the calendar description on <br> line break
                event_description_split = calendar_slice['description'].split('<br>')
                # split the event description first line (the hyperlink data) 
                event_description_line_one_split = event_description_split[0].split('"')
                # get the hyperlink available on the calendar event
                event_hyperlink = event_description_line_one_split[1]
                # get the calculated hyperlink from the shipping shceudle 
                ss_hyperlink = ss_item['GoogleSheetLink']
                # if the hyperlinks don't mathc then we need to update
                if ss_hyperlink != event_hyperlink:
                    # replace the hyperlink portion of the description with the shipping schedule hyperlink
                    event_description_line_one_split[1] = ss_hyperlink
                    # join the first line of the calendar description back together with the same splitting chracter
                    event_description_line_one_joined = '"'.join(event_description_line_one_split)
                    # change the first line of the event description
                    event_description_split[0] = event_description_line_one_joined
                    # join it back together with thte splitting character
                    new_desc = '<br>'.join(event_description_split)
                    # update the event with the new description & all the other same data
                    update_event(calendar_id = cal_id, 
                                event_id = calendar_slice['id'], 
                                summary = calendar_slice['summary'],
                                date = calendar_slice['date'].date(),
                                description = new_desc,
                                color_id = get_color(calendar_slice))  
                    print('hyperlink updated: ' + calendar_slice[type_of_work])
                    
                    add_to_change_log_v2(shop = shop, 
                                         job = calendar_slice[type_of_work].split('-')[0],
                                         type_of_work = type_of_work,
                                         number = calendar_slice[type_of_work],
                                         action = 'Hyperlink updated',
                                         description = '"' + ss_hyperlink + '"')                    
                
                
                
        except Exception as e:
            error_name = 'Misc Work Type Calendar Updating Error'
            
            produce_error_file(exception_as_e = e,
                                shop = shop,
                                file_prefix = error_name)  
            if _SendEmails:
                send_error_notice_email(shop, error_name, e)                
        
    

#%% This is for LOT's only


for shop in shop_dict.keys():
    print(shop)
    # get the calendar id from the shop_dict
    cal_id = shop_dict[shop]['cal_id']
    # get the recipients from the shop dict
    recipients_list = shop_dict[shop]['Recipients']
    # get the list of lots and their dates created as an excel file
    lots_df =  get_lots_information(shop)
    
    
    
    
    try:
        # get the current df of events
        events_df = convert_events_dict_to_df(cal_id, 'LOT')
        
        ''' Do a little spring cleaner of the calendar '''
        # get the events with no description
        to_delete = events_df[events_df['description'].isna()]
        # get the earliest delivery date in LOTS_df
        earliest_lots_df_date = min(lots_df['Earliest Delivery'])
        # get the calendar events 2 months older than that event
        to_delete = to_delete.append(events_df[events_df['date'] < earliest_lots_df_date - datetime.timedelta(days=3*7)])
        
        duplicates_to_delete = events_df[['LOT', 'id']].groupby(['LOT']).count()
        # get the ones with > 1 entry
        duplicates_to_delete = duplicates_to_delete[duplicates_to_delete['id'] > 1]
        # add those duplicates to the list to be deleted
        to_delete = to_delete.append(events_df[events_df['LOT'].isin(list(duplicates_to_delete.index))])
               
        # delete each of those events
        for event_id in to_delete['id']:
            row = events_df[events_df['id'] == event_id].squeeze()
            print('Spring cleaning of: ' + row['summary'])
            # delete the event
            delete_singular_event(cal_id, event_id)
            # pause 
            time.sleep(1)
        ''' end of the spring cleaning '''
        
        # get the current df of events after cleaning
        events_df = convert_events_dict_to_df(cal_id, 'LOT')
        # get the current timestamp as isoformat for the descipriont appendage
        now_timestamp = '(' + datetime.datetime.now().strftime('%m/%d/%Y %H:%M') + ')'  
        # first find the lots that are in the lots_df but not in the calendar
        missing_lots = np.setdiff1d(lots_df['LOT'], events_df['LOT'])
        if missing_lots.shape[0] == 0:
            print('No LOTS missing from the calendar after considerations')
        else:
            print('There are {} missing LOT(s) from the calendar after considerations'.format(missing_lots.shape[0]))
        
        # create the event for each lot that is missing
        for lot in missing_lots:
            # create the event name
            cal_name = 'LOT ' + lot
            # get the data from the 
            this_lot = lots_df[lots_df['LOT'] == lot].squeeze()
            # get the delivery date as a datetime.date()
            date = this_lot['Earliest Delivery'].date()
            # create a string that is the hyperlink display field
            sequences_str = 'Sequences: ' + this_lot['Sequences']
            if ',' in this_lot['Sequences']:
                sequences_str += ' (link goes to earliest delivery)'
            # generate the first line as the url to the first deliverable
            desc_first_line = '<a href="' + this_lot['URL'] + '">' + sequences_str + "</a>"
            # get the sequences as the description
            # desc = 'Sequences: ' + this_lot['Sequences']
            desc = desc_first_line 
            # desc += '<br>' + this_lot['URL']
            # include a tag for when the LOT was added to the calendar
            desc += '<br>' + now_timestamp + ' Added to Calendar with delivery date: ' + date.strftime('%m/%d/%Y')
            # get the comment for the lot, if there is any
            comment = this_lot['Comment']
            # if the comment is not na, then append the comment to the description
            if not pd.isna(comment):
                desc = desc + '<br>' + comment
            
            # create the event
            create_event(cal_id, cal_name, date, desc, color_id=11)
            # create a string of details for the change file
            change_details = cal_name + '\n' + date.isoformat() + '\n' + desc
            # create a file to denote the creation 
            produce_change_file(change_details, shop, lot, 'LOT Added')
            add_to_change_log_v2(shop = shop, 
                                job = lot.split('-')[0],
                                type_of_work = 'LOT',
                                number = lot,
                                action = 'New Calendar Event',
                                description = '')
            # wait incase of API limits
            time.sleep(1)
            # convert that to a dataframe
            new_events_df = convert_events_dict_to_df(cal_id, 'LOT')
            # get the df of the new event
            new_lot_event = new_events_df[new_events_df['summary'] == cal_name]
            if new_lot_event.shape[0] == 1:
                new_lot_event = new_lot_event.squeeze()
            else:
                print('THIS LOT HAS "MULTIPLE CALENDAR EVENTS: ', cal_name)
                
            # get the html link to that lot
            new_events_html_link = new_lot_event.squeeze()['htmlLink']
            if _SendEmails:
                # send an email out with the new lot information
                send_new_work_type_notice_email(shop, 'LOT', lot, date.strftime('%m/%d/%Y'), desc, 
                                      new_events_html_link, recipients_list)
    
        
        events_df = convert_events_dict_to_df(cal_id, 'LOT')
        
        ''' check for discrepancies in the sequences '''
        
        # get the sequences as found in the calendar, but only get the first line of the description
        cal_seqs = events_df['description'].str.split('<a', expand=True)[1]
        cal_seqs = cal_seqs.str.split('</a', expand=True)[0]
        cal_seqs = cal_seqs.str.split('>', expand=True)[1]
        cal_seqs = cal_seqs.str.split('(', expand=True)[0]
        # drop the string 'Sequences: '
        cal_seqs = cal_seqs.str[11:].str.strip()
        # set the index of the series to be the Lot name
        cal_seqs.index = events_df['LOT']
        # rename the series to be calendar sequences
        cal_seqs = cal_seqs.rename('cal seq').to_frame()
        # get the sequences of the lots from the production schedule
        ps_seqs = pd.DataFrame(data=list(lots_df['Sequences']),
                               index=lots_df['LOT'],
                               columns=['ps seq'])
        # join the 2 dfs
        updated_seqs = cal_seqs.join(ps_seqs)
        # compare the sequence strings of the lots
        updated_seqs = updated_seqs[updated_seqs['ps seq'] != updated_seqs['cal seq']]
        # get rid of any nan values when old lots are in the calendar
        updated_seqs = updated_seqs[~updated_seqs['ps seq'].isna()]
        
        
        if updated_seqs.shape[0] == 0:
            print('No LOTS need the sequences updated')
        for lot in updated_seqs.index:
            # get the current event information for that lot
            calendar_slice = events_df[events_df['LOT'] == lot].squeeze()
            # get the lots_df info for this lot
            this_lot = lots_df[lots_df['LOT'] == lot].squeeze()
            # the new sequences come from the production schedule
            new_seqs = this_lot['Sequences']
            # get the current calendar description
            cal_desc_str = calendar_slice['description']
            ''' FIxing the URL here'''
            # generate the hyperlink text
            sequences_str = 'Sequences: ' + new_seqs
            if ',' in new_seqs:
                sequences_str += ' (link goes to earliest delivery)'
            
            desc_first_line = '<a href="' + this_lot['URL'] + '">' + sequences_str + "</a>"
            
            
            # in case the calendar event has no description
            if type(cal_desc_str) == str:
                try:
                    # split the description by new line markers
                    cal_desc_list = cal_desc_str.split('<br>')
                    # save the string of old sequences
                    # old_seqs = cal_desc_list[0][11:]
                    old_seqs = [i for i in cal_desc_list if '>Sequences:' in i]
                    old_seqs = old_seqs[0].split('Sequences:')[1]
                    if '(' in old_seqs:
                        old_seqs = old_seqs[:old_seqs.find('(')]
                    old_seqs = old_seqs.strip(' ')
                    # replace the first item in the list version of the description
                    cal_desc_list[0] = desc_first_line
                    # join the list version with line breaks in between
                    new_desc = '<br>'.join(cal_desc_list)
                    # the string that gets added to the description so that you can see a history of the change
                    append_to_desc = now_timestamp + ' Sequences changed from [' + old_seqs +'] to ['+ new_seqs + ']'
                    # add the history line to the description
                    new_desc += '<br>' + append_to_desc
                except:
                    new_desc = desc_first_line
            else:
                date = this_lot['Earliest Delivery'].date()
                # get the sequences as the description
                new_desc = desc_first_line             
                new_desc += '(00/00/0000 00:00) Added to Calendar with delivery date: ' + date.strftime('%m/%d/%Y')
                new_desc += now_timestamp + ' Sequences are: [' + new_seqs + ']'
           
            
           
            
           # verify the delivery date from the production schedule 
            ps_date = lots_df[lots_df['LOT'] == lot].squeeze()['Earliest Delivery'].date()
            # get the calendar date
            cal_date = calendar_slice['date'].date()
            # if the calendar & production dates no longer line up, add a message to the description
            if ps_date != cal_date:
                # convert production schedule date to string format
                ps_date_str = ps_date.strftime('%m/%d/%Y')
                # convert the calendar date to a string format
                cal_date_str = cal_date.strftime('%m/%d/%Y')
                # create the new line message to add to the description
                append_to_desc = now_timestamp + ' Sequence change caused date change from ' + cal_date_str + ' to ' + ps_date_str
                # add to the description
                new_desc += '<br>' + append_to_desc
            
            
            
            # udpate the event with the new date (ps_date) & desc (new_desc), maintain everything else 
            update_event(calendar_id = cal_id, 
                         event_id = calendar_slice['id'], 
                         summary = calendar_slice['summary'],
                         date = ps_date,
                         description = new_desc,
                         color_id = get_color(calendar_slice))        
            
            # create a string of details for the change file
            change_details = calendar_slice['id'] + '\n' + calendar_slice['summary'] + '\n' + new_desc
            # create a file to denote the change of dates
            produce_change_file(change_details, shop, lot, 'Sequence Change')
            add_to_change_log_v2(shop = shop, 
                                job = lot.split('-')[0],
                                type_of_work = 'LOT',
                                number = lot,
                                action = 'Sequence Change',
                                description = 'old sequences -> ' + new_seqs.replace(',','|'))
            
            if _SendEmails:
                # send out an email regarding the change in dates
                send_seq_change_notice_email(shop, 'LOT', lot, ps_date.strftime('%m/%d/%Y'), new_desc, 
                                          calendar_slice['htmlLink'], recipients_list)
            
            time.sleep(1)            
        
        
        
        
        ''' check for discrepancies in the dates '''
        events_df = convert_events_dict_to_df(cal_id, 'LOT')

        # then find the lots that the dates do not match between events_df & lots_df
        # create a df that is just the lot & dates from the calendar
        cal_dates = pd.DataFrame(data=list(events_df['date']), 
                         index=events_df['LOT'], 
                         columns=['cal date'])
        # create a df that is just the lot & dates from the production schedule
        ps_dates = pd.DataFrame(data=list(lots_df['Earliest Delivery']), 
                          index=lots_df['LOT'], 
                          columns=['ps date'])
        # join the production schedule dates onto the calendar df
        updated_dates = cal_dates.join(ps_dates)
        # only keep rows that the calendar date and production schedule date do not align
        updated_dates = updated_dates[updated_dates['cal date'] != updated_dates['ps date']]
        # get rid of any NaT dates from the production schedule -> this happens after the sequences are completed
        updated_dates = updated_dates[~updated_dates['ps date'].isna()]
        
        if updated_dates.shape[0] == 0:
            print('No LOTS need the dates updated')
        else:
            print('{} LOT(s) need their dates udpated'.format(updated_dates.shape[0]))
        # go thru each lot in the updated_lots df
        for lot in updated_dates.index:
            # get the current event information for that lot
            calendar_slice = events_df[events_df['LOT'] == lot].squeeze()
            # get the current calendar date as iso-format
            cal_date_str = calendar_slice['date'].date().strftime('%m/%d/%Y')
            # get the date from the lots_df AKA the production schedule
            ps_date = lots_df[lots_df['LOT'] == lot].squeeze()['Earliest Delivery'].date()
            # conver the production schedule date to a iso-format
            ps_date_str = ps_date.strftime('%m/%d/%Y')

            # the string that gets added to the description so that you can see a history of the change
            append_to_desc = now_timestamp + ' Delivery date changed from ' + cal_date_str + ' to ' + ps_date_str            
            # get the description that is currently in the calendar event
            calendar_desc = calendar_slice['description']
            
            this_lot = lots_df[lots_df['LOT'] == lot].squeeze()
            # get the sequences for the lot
            seqs = this_lot['Sequences']
            # generate the hyperlink text
            sequences_str = 'Sequences: ' + seqs
            if ',' in seqs:
                sequences_str += ' (link goes to earliest delivery)'
            # generate the html for hyperlink
            desc_first_line = '<a href="' + this_lot['URL'] + '">' + sequences_str + "</a>"            
            # split the description based on the linebreak
            calendar_desc_list = calendar_desc.split('<br>')
            # change the hyperlink
            calendar_desc_list[0] = desc_first_line
            # join the description back with linebreaks
            calendar_desc = '<br>'.join(calendar_desc_list)
            
            # if there is no descirption (unlikely) then just set the append_to_desc as the desc
            if pd.isna(calendar_desc):
                new_desc = desc_first_line + '<br>' + append_to_desc
            # if there is a description already on the event, then append the append_to_desc
            else:
                new_desc = calendar_desc + '<br>' + append_to_desc
            
            
            
            # udpate the event with the new date (ps_date) & desc (new_desc), maintain everything else 
            update_event(calendar_id = cal_id, 
                         event_id = calendar_slice['id'], 
                         summary = calendar_slice['summary'],
                         date = ps_date,
                         description = new_desc,
                         color_id = get_color(calendar_slice))        
            
            
            # create a string of details for the change file
            change_details = calendar_slice['id'] + '\n' + calendar_slice['summary'] + '\n' + new_desc
            # create a file to denote the change of dates
            produce_change_file(change_details, shop, lot, 'Date Change')
            add_to_change_log_v2(shop = shop, 
                                job = lot.split('-')[0],
                                type_of_work = 'LOT',
                                number = lot,
                                action = 'Change Event Date',
                                description = cal_date_str + ' -> ' + ps_date_str)            
            
            if _SendEmails:
                # send out an email regarding the change in dates
                send_date_change_notice_email(shop, 'LOT', lot, ps_date.strftime('%m/%d/%Y'), new_desc, 
                                          calendar_slice['htmlLink'], recipients_list)
            
            time.sleep(1)
            
        
        ''' update the hyperlinks '''
        events_df = convert_events_dict_to_df(cal_id, 'LOT')
        for i in events_df.index:
            # get the calendar data from events_df
            calendar_slice = events_df.loc[i].squeeze()
            # get the information from the shipping schedule
            ss_item = lots_df[lots_df['LOT'] == calendar_slice['LOT']].squeeze()
            if not ss_item.shape[0]:
                continue
            # split the calendar description on <br> line break
            event_description_split = calendar_slice['description'].split('<br>')
            # split the event description first line (the hyperlink data) 
            event_description_line_one_split = event_description_split[0].split('"')
            # get the hyperlink available on the calendar event
            event_hyperlink = event_description_line_one_split[1]
            # get the calculated hyperlink from the shipping shceudle 
            ss_hyperlink = ss_item['URL']
            # if the hyperlinks don't mathc then we need to update
            if ss_hyperlink != event_hyperlink:
                # replace the hyperlink portion of the description with the shipping schedule hyperlink
                event_description_line_one_split[1] = ss_hyperlink
                # join the first line of the calendar description back together with the same splitting chracter
                event_description_line_one_joined = '"'.join(event_description_line_one_split)
                # change the first line of the event description
                event_description_split[0] = event_description_line_one_joined
                # join it back together with thte splitting character
                new_desc = '<br>'.join(event_description_split)
                # update the event with the new description & all the other same data
                update_event(calendar_id = cal_id, 
                            event_id = calendar_slice['id'], 
                            summary = calendar_slice['summary'],
                            date = calendar_slice['date'].date(),
                            description = new_desc,
                            color_id = get_color(calendar_slice))  
                print('hyperlink updated: ' + calendar_slice['LOT'])
                
                add_to_change_log_v2(shop = shop, 
                                    job = calendar_slice['LOT'].split('-')[0],
                                    type_of_work = 'LOT',
                                    number = calendar_slice['LOT'],
                                    action = 'Hyperlink Updated',
                                    description = '"' + ss_hyperlink + '"')   
                time.sleep(1)
    
   
    
    except Exception as e:
        error_name = 'LOT Calendar Updating Error'
        
        produce_error_file(exception_as_e = e,
                           shop = shop,
                           file_prefix = error_name)  
        if _SendEmails:
            send_error_notice_email(shop, error_name, e)
        
        

#%% manila calendar

manila_calendar_function()


#%%

# all_recipients = ['cwilson@crystalsteel.net']

# the_datetime = datetime.datetime.now()
# if the_datetime.hour == 23:
    
#     todays_updates = changes_and_new_work(day_as_date = the_datetime.date())
    
#     daily_changes_new_work_email(days_date_str = the_datetime.date().strftime('%Y-%m-%d'),
#                                   new_and_changes_dict = todays_updates,
#                                   recipients_list = all_recipients)
    
    
    
#     todays_errors = daily_errors_summary()
    
#     daily_errors_email(days_date_str = the_datetime.date().strftime('%Y-%m-%d'),
#                        errors_dataframe = todays_errors,
#                        recipients_list = ['csf_pm@crystalsteel.net', 'cwilson@crystalsteel.net', 'lcolon@crystalsteel.net'])
