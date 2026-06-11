# -*- coding: utf-8 -*-
"""
Created on Thu Jun 11 09:02:00 2026

@author: Netadmin
"""


from pathlib import Path
import os
    
def returnPrimePointCredentials():
    filename = Path(os.getcwd()) / 'PrimePointCredentials.txt'
    
    
    try: 
        with open(filename) as file:
            lines = [line.rstrip() for line in file]
    except Exception as e:
        raise e 
    
    username = lines[0].split(':')[1].strip()
    password = lines[1].split(':')[1].strip()
    
    return {'username': username, 'password': password}