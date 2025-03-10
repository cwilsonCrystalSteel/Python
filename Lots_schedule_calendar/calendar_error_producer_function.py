# -*- coding: utf-8 -*-
"""
Created on Thu Jun 17 09:46:02 2021

@author: CWilson
"""
import datetime
import os
from pathlib import Path

def produce_error_file(exception_as_e, shop, file_prefix):
    error_folder = Path(os.getcwd()) / 'Lots_schedule_calendar' / 'Error_Logs'
    if not os.path.exists(error_folder):
        os.makedirs(error_folder)
    # the current tiemstamp for the file
    error_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # the files name
    file_name = error_folder / (shop + ' ' + file_prefix + ' ' + error_date + '.txt')
    # open the file
    file = open(file_name, 'w')
    # write the shop name in the file
    file.write(shop + '\n')
    # write the error information to the file
    file.write(str(exception_as_e))
    # close the file
    file.close()