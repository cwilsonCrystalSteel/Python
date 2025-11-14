# -*- coding: utf-8 -*-
"""
Created on Thu May  8 11:21:23 2025

@author: Netadmin
"""

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



def email_pdf_report(pdf_filepath, month_name, year, dfs_to_fix_dict, state, recipients, xlsx_filepath = None):
    service = get_email_service()
    # import shutil
    print(f'Sending {state} {month_name} {year} Fitter & Welder Stats Report email to: ', recipients)
    # the reporting email address
    my_email = 'csmreporting@crystalsteel.net'
    # the directory to save the csv files as a copy to
    
    # Create the first line of the email
    email_start = f"\n<p>Fitter & Welder Stats for {month_name} {year}</p>\n"
    # create the line break and "Breakdown" 
    email_middle = '\n\n<p style="font-size:16px;"><u>See Attached PDF For Report</u></p><br>\n'
    
    email_msg = email_start + email_middle
    
    if dfs_to_fix_dict:
        fix_header = '<hr><p style="font-size:16px;">The below tables are the different groupings of records that have incorrectly labeled Fitter/Welder Employee ID Numbers</p><br>'
        link_warning_msg = "<b><u>THESE URLs ARE SUBJECT TO CHANGE IF TO MUCH TIME HAS PASSED</b></u><br>\n--> reply to this email and cwilson@crystalsteel.net for an updated email"
        email_msg += fix_header + link_warning_msg
    
    
    for table in dfs_to_fix_dict:
        table_text = f"\n<p>Bad Entries Caused By: <b>{table}</b></p>\n"
        print(table)
        df = dfs_to_fix_dict[table].copy()
        # set to object type first to prevent error/warning
        df.iloc[:, 2] = df.iloc[:, 2].astype("object")
        df.iloc[:, 3] = df.iloc[:, 3].astype("object")
        # fitter/welder id columns
        df.iloc[:, 2] = df.iloc[:, 2].apply(lambda x: x if isinstance(x, str) else (f"{int(x)}" if pd.notnull(x) else ""))
        # job number
        df.iloc[:, 3] = df.iloc[:, 3].apply(lambda x: x if isinstance(x, str) else (f"{int(x)}" if pd.notnull(x) else ""))

        table_text += df.to_html(index=False, escape=False)
        table_text += "\n<br>\n"
    
        email_msg += table_text
    
    
    
    # little note regarding who to contact about the reports
    email_end = "\n<br>\n<p>Please email cwilson@crystalsteel.net for questions regarding this information</p>\n"   
    # combine all the HTML strings to form the message
    email_msg +=  email_end
    # create new message
    msg = MIMEMultipart("html")
    # create the subject line
    msg['Subject'] = f'{state} Fitter & Welder Stats Report for {month_name} {year}'
    # create the sender
    msg['From'] = f'{state} Fitter & Welder Stats Report <{my_email}>'
    # create the receiver
    msg['To'] = ', '.join(recipients)
    # create the body of the message
    part1 = MIMEText(email_msg, 'html')
    # attach the bod to the message
    msg.attach(part1)

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(pdf_filepath, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(pdf_filepath)))
    msg.attach(part)
    
    if xlsx_filepath is not None:
        part1 = MIMEBase('application', "octet-stream")
        part1.set_payload(open(xlsx_filepath, "rb").read())
        encoders.encode_base64(part1)
        part1.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(xlsx_filepath)))
        msg.attach(part1)

    
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