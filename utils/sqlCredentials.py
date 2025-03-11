# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 16:30:24 2024

@author: CWilson
"""
import os
from pathlib import Path

def returnSqlCredentials():
    credsPath = Path(os.getcwd()) / 'sqlCredentials.txt'
    with open(credsPath) as f:
        out = f.readlines()
        
    out = [i.replace('\n','') for i in out]
    outdict = {i.split(':')[0]:i.split(':')[1] for i in out}
    return outdict