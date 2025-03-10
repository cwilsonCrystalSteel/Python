# -*- coding: utf-8 -*-
"""
Created on Wed Jun 16 09:25:56 2021

@author: CWilson
"""

import datetime
import pickle
import os
from pathlib import Path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

CREDENTIALS_FILE = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'lots_calendar.json'

PICKLE_PATH = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'calendar_token.json'


def get_calendar_service():
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

   service = build('calendar', 'v3', credentials=creds)
   
   return service