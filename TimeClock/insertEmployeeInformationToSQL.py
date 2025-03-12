# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 15:39:29 2024

@author: CWilson
"""

''' Daily 
 of the employee information '''


from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
import pandas as pd
from TimeClock.TimeClockNavigation import TimeClockBase
from utils.initSQLConnectionEngine import yield_SQL_engine
from pathlib import Path

download_path = Path.home() / 'Downloads' / 'EmployeeInformation'
engine = yield_SQL_engine()

def print_count_results(schema, engine, suffix_text):
    with engine.connect() as connection:
        result = connection.execute(text(f"select count(*) from {schema}.employeeinformation"))
        for row in result:
            continue
    
    print(f"There are {row[0]} rows in {schema}.employeeinformation {suffix_text}")
    
    
def print_terminated_count_results(engine, terminated, suffix_text):
    terminated_str = 'TRUE' if terminated else 'FALSE'
    with engine.connect() as connection:
        result = connection.execute(text(f"select count(*) from dbo.employeeinformation where terminated={terminated_str}"))
        for row in result:
            continue
        
    print(f"There are {row[0]} rows in dbo.employeeinformation where terminated={terminated_str} {suffix_text}")
        
    

def import_employee_information_to_SQL():

    print(f'{"*"*50}\nBegining to import employee information to SQL\n{"*"*50}')
    
    x = TimeClockBase(download_path, headless=True)     
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
    
    x.kill()
    
    print_count_results('live', engine, 'before truncating')
    
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.employeeinformation'''))
    session.commit()
    session.close()
    
    
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

       

def determine_terminated_employees():
    print(f'{"*"*50}\nBegining to determine terminated employees into SQL\n{"*"*50}')
    
    # Get the employee information for all employees
    x = TimeClockBase(download_path, headless=True)     
    x.verbosity=0
    x.startupBrowser()
    x.tryLogin()
    x.openTabularMenu()
    x.searchFromTabularMenu('export')
    x.clickTabularMenuSearchResults('Tools > Export')
    try:
        x.employeeLocationFinale()
        filepath_all = x.retrieveDownloadedFile(10, '*.csv', 'Employee Information')
        print(filepath_all)
    except Exception as e:
        print(f'Could not complete download because {e}')
        
    ei_all = pd.read_csv(filepath_all)
    
    x.kill()
    
    # get the employee information for terminated employees
    x = TimeClockBase(download_path, headless=True)   
    x.verbosity=0
    x.startupBrowser()
    x.tryLogin()
    x.openTabularMenu()
    x.searchFromTabularMenu('export')
    x.clickTabularMenuSearchResults('Tools > Export')
    try:
        x.employeeLocationFinale(exclude_terminated=True)
        filepath_notTerminated = x.retrieveDownloadedFile(10, '*.csv', 'Employee Information')
        print(filepath_notTerminated)
    except Exception as e:
        print(f'Could not complete download because {e}')
        
    ei_notTerminated = pd.read_csv(filepath_notTerminated)
    
    x.kill()
    
    ei_terminated = pd.merge(ei_all, ei_notTerminated, left_on='<NUMBER>', right_on='<NUMBER>', suffixes=('','_notTerminated'), how='left')
    ei_terminated = ei_terminated[ei_terminated['<FIRSTNAME>_notTerminated'].isna()]
    # need to convert the headers back to how they should be 
    ei_terminated = ei_terminated.drop(columns=[i for i in ei_terminated.columns if '_notTerminated' in i])
    
    
    
    # truncate live.employeeinformation
    # insert ei_terminated into live.employeeInformation
    # use merge proc update_terminatedemployees
        # inner join live.employeeinformation to dbo.employeeinformation
        # update values of terminated
    print('RUNNING UPDATES FOR TERMINATED EMPLOYEES')
    print_count_results('live', engine, 'before truncating')
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.employeeinformation'''))
    session.commit()
    session.close()
    
    print_count_results('live', engine, 'before importing')
    
    ei_terminated.to_sql('employeeinformation', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results('live', engine, 'after importing')
    
    print_terminated_count_results(engine,True, 'before updating IS terminated')
    print_terminated_count_results(engine,False, 'before updating IS terminated')
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call dbo.update_employeeinformation_isterminated()")
    connection.commit()
    connection.close()
    print_terminated_count_results(engine,True, 'after updating IS terminated')
    print_terminated_count_results(engine,False, 'after updating IS terminated')
    
    # double check length of live table after merge proc --> should be 0
    print_count_results('live', engine, 'after merge proc')
    
    
    
    
    print('RUNNING UPDATES FOR NON-TERMINATED EMPLOYEES')
    print_count_results('live', engine, 'before truncating')
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.employeeinformation'''))
    session.commit()
    session.close()
    
    print_count_results('live', engine, 'before importing')
    
    ei_notTerminated.to_sql('employeeinformation', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results('live', engine, 'after importing')
    
    print_terminated_count_results(engine,True, 'before updating IS terminated')
    print_terminated_count_results(engine,False, 'before updating IS terminated')
    connection = engine.raw_connection()
    cursor = connection.cursor()
    cursor.execute("call dbo.update_employeeinformation_isnnotterminated()")
    connection.commit()
    connection.close()
    print_terminated_count_results(engine,True, 'after updating IS terminated')
    print_terminated_count_results(engine,False, 'after updating IS terminated')
    
    # double check length of live table after merge proc --> should be 0
    print_count_results('live', engine, 'after merge proc')
    