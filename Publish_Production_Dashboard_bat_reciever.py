# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 15:25:56 2021

@author: CWilson
"""

import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
from Publish_Production_Dashboard import publish_dashboard



counter_file = 'C:\\users\\cwilson\\documents\\python\\Production_Dashboard_temp_files\\Production_dashboard_counter.txt'
f = open(counter_file, 'r')
count = int(f.readlines()[-1])
f = open(counter_file, 'w')
f.write(str(count + 1))
f.close()

# no amtter what, run the daily
publish_dashboard('Daily')

# if it has ran 50 times, run yearly
if not count % 50:
    publish_dashboard('Yearly')
    # if it has run 20 times, run monthly
elif not count % 20:
    publish_dashboard('Monthly')
    # if it has run 8 times, run weekly
elif not count % 5:
    publish_dashboard('Weekly')
    