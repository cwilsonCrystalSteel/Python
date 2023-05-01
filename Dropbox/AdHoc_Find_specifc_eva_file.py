# -*- coding: utf-8 -*-
"""
Created on Fri Apr 28 07:34:37 2023

@author: CWilson
"""

import os


prereq = '//192.168.50.9//Dropbox_(CSF)//'
os.scandir(prereq)

# this is the base dir from the X:\\
base_dir = 'X:\\production control\\EVA REPORTS FOR THE DAY\\'

f = []
for path, subdirs, files in os.walk(base_dir):
    for name in files:
        f.append(os.path.join(path, name))
        
        
j = [i for i in f if '2218' in i]
k = [i for i in j if '103' in i]

l = [i for i in f if '103-' in i]
