# -*- coding: utf-8 -*-
"""
Created on Wed Jun 16 10:23:24 2021

@author: CWilson
"""

import datetime
from cal_setup import get_calendar_service

def main():
   service = get_calendar_service()
   # Call the Calendar API
   now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
   print('Getting List o 10 events')
   events_result = service.events().list(
       calendarId='c_uqopq4705q7o473uveoah08h7o@group.calendar.google.com', timeMin=now,
       maxResults=10, singleEvents=True,
       orderBy='startTime').execute()
   events = events_result.get('items', [])

   if not events:
       print('No upcoming events found.')
   for event in events:
       start = event['start'].get('dateTime', event['start'].get('date'))
       print(start, event['summary'])

if __name__ == '__main__':
    main()