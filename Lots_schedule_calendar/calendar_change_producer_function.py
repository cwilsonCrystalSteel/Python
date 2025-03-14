# -*- coding: utf-8 -*-
"""
Created on Thu Jun 17 11:54:48 2021

@author: CWilson
"""

import datetime
import os
from pathlib import Path

def produce_change_file(change_details, shop, name, file_prefix):
    change_folder = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'Change_logs'
    if not os.path.exists(change_folder):
        os.makedirs(change_folder)
    # the current tiemstamp for the file
    change_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # the files name
    file_name = change_folder / (file_prefix + ' ' + change_date + '.txt')
    # open the file
    file = open(file_name, 'w')
    # write the shop name in the file
    file.write(shop + '\n')
    # write the lot name in the file
    file.write(name + '\n')
    # write the error information to the file
    file.write(change_details)
    # close the file
    file.close()