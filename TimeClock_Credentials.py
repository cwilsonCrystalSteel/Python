# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 09:35:59 2023

@author: CWilson
"""

    
def returnTimeClockCredentials():
    filename = 'C:\\Users\\cwilson\\Documents\\Python\\TimeClockCredentials.txt'
    
    with open(filename) as file:
        lines = [line.rstrip() for line in file]
    
    username = lines[0].split(':')[1].strip()
    password = lines[1].split(':')[1].strip()
    
    return {'username': username, 'password': password}