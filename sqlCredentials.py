# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 16:30:24 2024

@author: CWilson
"""

def returnSqlCredentials():
    with open(r'c:\users\cwilson\documents\python\sqlCredentials.txt') as f:
        out = f.readlines()
        
    out = [i.replace('\n','') for i in out]
    outdict = {i.split(':')[0]:i.split(':')[1] for i in out}
    return outdict