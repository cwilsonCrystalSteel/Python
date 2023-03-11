# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:36:21 2023

@author: CWilson
"""

from production_dashboards_google_credentials import init_google_sheet

# this one will send TimeClock & Fablisting data to google sheet



google_sheet_info = {'sheet_key':'1RZKV2-jt5YOFJNKM8EJMnmAmgRM1LnA9R2-Yws2XQEs',
                     'json_file':'C:\\Users\\cwilson\\Documents\\Python\\production-dashboard-other-890ed2bf828b.json'
                     }
shop = 'CSM'
sh = init_google_sheet(google_sheet_info['sheeet_key'], google_sheet_info['json_file'])
worksheet = sh.worksheet(shop)
