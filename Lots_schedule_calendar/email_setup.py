# -*- coding: utf-8 -*-
"""
Created on Thu Jun  2 09:56:04 2022

@author: CWilson
"""



import datetime
import pickle
import os
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


SCOPES = ['https://mail.google.com/']

CREDENTIALS_FILE = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'lots_calendar.json'

PICKLE_PATH = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'email_token.pickle'



def get_email_service():
   creds = None
   # The file token.pickle stores the user's access and refresh tokens, and is
   # created automatically when the authorization flow completes for the first
   # time.
   
   
   if os.path.exists(PICKLE_PATH):
       with open(PICKLE_PATH, 'rb') as token:
           creds = pickle.load(token)
   # If there are no (valid) credentials available, let the user log in.
   if not creds or not creds.valid:
       if creds and creds.expired and creds.refresh_token:
           creds.refresh(Request())
       else:
           flow = InstalledAppFlow.from_client_secrets_file(
               CREDENTIALS_FILE, SCOPES)
           creds = flow.run_local_server(port=0)

       # Save the credentials for the next run
       with open(PICKLE_PATH, 'wb') as token:
           pickle.dump(creds, token)

   service = build('gmail', 'v1', credentials=creds)
   
   return service