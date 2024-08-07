# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 15:39:29 2024

@author: CWilson
"""

''' Daily download of the employee information '''


from sqlalchemy import text
import pandas as pd
from TimeClockNavigation import TimeClockBase
from initSQLConnectionEngine import yield_SQL_engine
from Gather_data_for_timeclock_based_email_reports import get_ei_csv_downloaded



engine = yield_SQL_engine()

def print_count_results(schema, engine, suffix_text):
    with engine.connect() as connection:
        result = connection.execute(text(f"select count(*) from {schema}.employeeinformation"))
        for row in result:
            continue
    
    print(f"There are {row[0]} rows in {schema}.employeeinformation {suffix_text}")
    

def import_employee_information_to_SQL():

    

    x = TimeClockBase("C:\\users\\cwilson\\downloads\\EmployeeInformation\\", headless=True)     
    x.startupBrowser()
    x.tryLogin()
    x.openTabularMenu()
    x.searchFromTabularMenu('export')
    x.clickTabularMenuSearchResults('Tools > Export')
    try:
        x.employeeLocationFinale()
        filepath = x.retrieveDownloadedFile(10, '*.csv', 'Employee Information')
        print(filepath)
    except Exception as e:
        print(f'Could not complete download because {e}')
        
    ei = pd.read_csv(filepath)
    
    # get count of table before insert --> should be 0
    print_count_results('live', engine, 'before importing')
    print_count_results('dbo', engine, 'before merge proc')
    
    ei.to_sql('employeeinformation', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results('live', engine, 'after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call dbo.merge_employeeinformation()")
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results('dbo', engine, 'after merge proc') 
    print_count_results('live', engine, 'after merge proc')

       

    