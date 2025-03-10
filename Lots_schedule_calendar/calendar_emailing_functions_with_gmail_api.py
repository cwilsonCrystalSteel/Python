# -*- coding: utf-8 -*-
"""
Created on Fri Jun 25 14:39:06 2021

@author: CWilson
"""

import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import google.auth
import base64
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Lots_schedule_calendar.email_setup import get_email_service








# # the reporting email address
# my_email = 'lotscalendar.csf@gmail.com'
# # IDGAF about account security !!!
# password = "Memphis2020!"

my_email = 'csmreporting@crystalsteel.net'

def send_new_work_type_notice_email(shop, type_of_work, name, date, desc, event_url_link, recipients_list):
    service = get_email_service()
    # define a line break in HTML to save time and space
    br = "<br></br>\n"
    # replace any \n line breaks in the description with a HTML line break
    desc = desc.replace('\n', br)
    
    print(shop + ' has a new ' + type_of_work)
    print(type_of_work + ' ' + name + ' delivers on:\t' + date)

    # first email line
    email_preface = "\n<p>A new " + type_of_work + " has been added to the calendar for " + shop+ "\n"
    # body line 1
    email_1 = 2*br + ' ' + type_of_work + ' ' + name + " is scheduled to deliver on " + date + "\n"
    # body line 2
    email_2 = 2*br + "<u>Event Description:</u><br></br>\n" + desc + "\n"
    # link to the calendar event
    email_3 = 2*br + "<u>Calendar Event Link:</u><br></br>\n" + event_url_link + "</p>\n"
    # little note regarding who to contact about the reports
    email_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # combine all the HTML strings to form the message
    email_msg = email_preface + email_1 + email_2 + email_3 + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = shop + ' New ' + type_of_work + ': ' + name
    # create the sender
    msg['From'] = shop + ' New ' + type_of_work + '<' + my_email +'>'
    # create the receiver
    msg['To'] = ", ".join(recipients_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')
    
    

        
        
def send_date_change_notice_email(shop, type_of_work, name, date, desc, event_url_link, recipients_list):
    service = get_email_service()
    # define a line break in HTML to save time and space
    br = "<br></br>\n"
    # replace any \n line breaks in the description with a HTML line break
    desc = desc.replace('\n', br)
    
    print(shop + ' has a ' + type_of_work + ' with a changed delivery date')
    print(name + ' now delivers on:\t' + date)
    # first email line
    email_preface = "\n<p>A change in delivery date has been detected for a " + shop + " " + type_of_work + "\n"
    # body line 1
    email_1 = 2*br + type_of_work + ' ' + name + " is now scheduled to deliver on " + date + "\n"
    # body line 2
    email_2 = 2*br + "<u>Event Description:</u><br></br>\n" + desc + "\n"
    # link to the calendar event
    email_3 = 2*br + "<u>Calendar Event Link:</u><br></br>\n" + event_url_link + "</p>\n"
    # little note regarding who to contact about the reports
    email_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # combine all the HTML strings to form the message
    email_msg = email_preface + email_1 + email_2 + email_3 + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = shop + ' Delivery Date change for: ' + type_of_work + ' ' + name
    # create the sender
    msg['From'] = shop + ' ' + type_of_work + ' Date change<' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(recipients_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')    
       
        

        
def send_seq_change_notice_email(shop, type_of_work, name, date, desc, event_url_link, recipients_list):
    service = get_email_service()
    # define a line break in HTML to save time and space
    br = "<br></br>\n"
    # replace any \n line breaks in the description with a HTML line break
    desc = desc.replace('\n', br)
    
    print(shop + ' has a ' + type_of_work + ' with changed Sequences')
    # first email line
    email_preface = "\n<p>A change in Sequences has been detected for a " + shop + " " + type_of_work + "\n"
    # body line 1
    email_1 = 2*br + type_of_work + ' ' + name + " is now scheduled to deliver on " + date + "\n"
    # body line 2
    email_2 = 2*br + "<u>Event Description:</u><br></br>\n" + desc + "\n"
    # link to the calendar event
    email_3 = 2*br + "<u>Calendar Event Link:</u><br></br>\n" + event_url_link + "</p>\n"
    # little note regarding who to contact about the reports
    email_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # combine all the HTML strings to form the message
    email_msg = email_preface + email_1 + email_2 + email_3 + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = shop + ' Sequence change for: ' + type_of_work + ' ' + name
    # create the sender
    msg['From'] = shop + ' Sequence change<' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(recipients_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')
    
   




def send_error_notice_email(shop, error_type, error_as_e, extra_recipient=None):
    service = get_email_service()
    
    lots_log_errors_email_recipients = ['cwilson@crystalsteel.net', 'jturner@crystalsteel.net']
    
    # lots_log_errors_email_recipients = ['cwilson@crystalsteel.net']
    
    
    print(error_type)
    # if there is only one email / the email is a string, add it to the list of recipients but as a list
    # if there is a list of extra people, just add list to tlist 
    if extra_recipient != None:
        if isinstance(extra_recipient, str):
            lots_log_errors_email_recipients += [extra_recipient]
        elif isinstance(extra_recipient, list):
            lots_log_errors_email_recipients += extra_recipient
        
    send_error_messages_to = ', '.join(lots_log_errors_email_recipients) 
    # send_error_messages_to = 'cwilson@crystalsteel.net'
    print('Sending errors to: {}'.format(send_error_messages_to))
    # define a line break in HTML to save time and space
    br = "<br></br>\n"
    # first email line
    email_preface = "\n<p>An error has occured for " + shop + "\n"
    # body line 1
    email_1 = 2*br + error_type + "\n"
    # body line 2
    email_2 = 2*br + str(error_as_e) + "</p>\n"
    # little note regarding who to contact about the reports
    email_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # combine all the HTML strings to form the message
    email_msg = email_preface + email_1 + email_2 + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = 'LOT CALENDAR ERROR'
    # create the sender
    msg['From'] = 'ERROR <' + my_email +'>'
    # create the receiver
    msg['To'] = send_error_messages_to
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')    
   
    
   


        
def daily_changes_new_work_email(days_date_str, new_and_changes_dict, recipients_list):
    service = get_email_service()
    # define a line break in HTML to save time and space
    br = "<br></br>\n"
    # replace any \n line breaks in the description with a HTML line break

    new_work = new_and_changes_dict['New']
    new_work = new_work.drop(columns=['today'])
    changes = new_and_changes_dict['Changes']
    changes = changes.drop(columns=['today'])
    
    new_work_html = new_work.to_html(col_space=100, justify='center', index=False)
    changes_html = changes.to_html(col_space=100, justify='center', index=False)

    # first email line
    email_1 = "\n<p>There are " + str(new_work.shape[0]) + " new entries to the Shipping Schedule</p>\n"
    email_2 = "\n<p>There are " + str(changes.shape[0]) + " changes to the Shipping Schedule\n"

    # little note regarding who to contact about the reports
    email_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # combine all the HTML strings to form the message
    email_msg = email_1 + new_work_html + email_2 + changes_html + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = days_date_str + ' Updates to Shipping Schedule'
    # create the sender
    msg['From'] = 'Shipping Schedule Updates<' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(recipients_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')
    



def daily_errors_email(days_date_str, errors_dataframe, recipients_list):
    service = get_email_service()
    # define a line break in HTML to save time and space
    br = "<br></br>\n"
    # replace any \n line breaks in the description with a HTML line break

    week_plus = errors_dataframe[errors_dataframe['Number of Days Error Present']>= 7].shape[0]
    
    errors_dataframe_html = errors_dataframe.to_html(col_space=100, justify='center', index=False)

    # first email line
    email_1 = "\n<p>There are " + str(errors_dataframe.shape[0]) + " errors in the Shipping Schedule</p>\n"
    email_2 = "\n<p>There are " + str(week_plus) + " errors that are more than a week old!</p>\n"
    email_3 = "\n<p>* Some errors may show up here because they were caught & recorded before being fixed, and may be fixed now *</p>\n"

    # little note regarding who to contact about the reports
    email_end = "\n<br></br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"    
    # combine all the HTML strings to form the message
    email_msg = email_1 + email_2 + email_3 + errors_dataframe_html  + email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = days_date_str + ' Shipping Schedule Errors'
    # create the sender
    msg['From'] = 'Shipping Schedule Errors<' + my_email +'>'
    # create the receiver
    msg['To'] = ', '.join(recipients_list)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)
    
    encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    create_message = {
        'raw': encoded_message
    }
    # pylint: disable=E1101
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')