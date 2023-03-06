# -*- coding: utf-8 -*-
"""
Created on Wed Jun 16 09:40:30 2021

@author: CWilson
"""

from cal_setup import get_calendar_service

def main():
   service = get_calendar_service()
   # Call the Calendar API
   print('Getting list of calendars')
   calendars_result = service.calendarList().list().execute()

   calendars = calendars_result.get('items', [])

   if not calendars:
       print('No calendars found.')
   for calendar in calendars:
       summary = calendar['summary']
       id = calendar['id']
       primary = "Primary" if calendar.get('primary') else ""
       print("%s\t%s\t%s" % (summary, id, primary))

if __name__ == '__main__':
   main()