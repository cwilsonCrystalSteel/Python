# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 11:37:02 2022

@author: CWilson
"""



from High_Indirect_Hours_Email_Report import emaIL_attendance_hours_report


weekstart = '09/04/2022'
state = 'TN'
filepath = 'c:\\users\\cwilson\\documents\\Productive_Employees_Hours_Worked_report\\week_by_week_hours_of_employees TN_formatted.xlsx'
state_recipients = {'TN': ['cwilson@crystalsteel.net'],
 'MD': ['cwilson@crystalsteel.net'],
 'DE': ['cwilson@crystalsteel.net']}
emaIL_attendance_hours_report(weekstart, state, filepath, state_recipients)