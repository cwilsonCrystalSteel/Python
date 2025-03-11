# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 09:35:59 2023

@author: CWilson
"""

from pathlib import Path
import os
    
def returnTimeClockCredentials():
    filename = Path(os.getcwd()) / 'TimeClockCredentials.txt'
    
    
    try: 
        with open(filename) as file:
            lines = [line.rstrip() for line in file]
    except Exception as e:
        raise e 
    
    username = lines[0].split(':')[1].strip()
    password = lines[1].split(':')[1].strip()
    
    return {'username': username, 'password': password}