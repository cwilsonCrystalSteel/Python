# -*- coding: utf-8 -*-
"""
Created on Fri Apr 28 07:34:37 2023

@author: CWilson
"""

import os
from pathlib import Path


 
possible_dir = ['Y:','X:','\\\\192.168.50.9\\Dropbox_(CSF)']
for ii in possible_dir:
    if os.path.exists(Path(ii)):
        base_dir = Path(ii) / 'production control' / 'EVA REPORTS FOR THE DAY'
        print(f'Using the drive: "{base_dir}"')
        break
    else:
        continue

f = []
for path, subdirs, files in os.walk(base_dir):
    for name in files:
        f.append(os.path.join(path, name))
        
        
j = [i for i in f if '2218' in i]
k = [i for i in j if '103' in i]

l = [i for i in f if '103-' in i]
