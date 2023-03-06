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

# the reporting email address
my_email = 'lotscalendar.csf@gmail.com'
# IDGAF about account security !!!
password = "Memphis2020!"

def send_new_work_type_notice_email(shop, type_of_work, name, date, desc, event_url_link, recipients_list):
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

    # for SSL
    port = 465         
    # initialize email stuff
    context = ssl.create_default_context()
    # start the email up 
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(my_email, password)
        server.send_message(msg, from_addr='lotscalendar.csf@gmail.com', to_addrs=recipients_list)
        # server.set_debuglevel(1)
        server.quit()
        
        
def send_date_change_notice_email(shop, type_of_work, name, date, desc, event_url_link, recipients_list):
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
    # for SSL
    port = 465         
    # initialize email stuff
    context = ssl.create_default_context()
    # start the email up 
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(my_email, password)
        server.send_message(msg, from_addr='lotscalendar.csf@gmail.com', to_addrs=recipients_list)
        # server.set_debuglevel(1)
        server.quit()     
        

        
def send_seq_change_notice_email(shop, type_of_work, name, date, desc, event_url_link, recipients_list):
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
    # for SSL
    port = 465         
    # initialize email stuff
    context = ssl.create_default_context()
    # start the email up 
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(my_email, password)
        server.send_message(msg, from_addr='lotscalendar.csf@gmail.com', to_addrs=recipients_list)
        # server.set_debuglevel(1)
        server.quit()  




def send_error_notice_email(shop, error_type, error_as_e, extra_recipient=None):
    print(error_type)
    send_error_messages_to = 'cwilson@crystalsteel.net, cwilson@crystalsteel.net, jturner@crystalsteel.net'
    if extra_recipient:
        send_error_messages_to += ', ' + extra_recipient
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
    # for SSL
    port = 465         
    # initialize email stuff
    context = ssl.create_default_context()
    # start the email up 
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(my_email, password)
        server.send_message(msg, from_addr='lotscalendar.csf@gmail.com', to_addrs=send_error_messages_to)
        # server.set_debuglevel(1)
        server.quit()  