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
from utils.sql_print_count_results import print_count_results, print_count_results_where
from pathlib import Path

download_path = Path.home() / 'Downloads' / 'EmployeeInformation'
engine = yield_SQL_engine()

table = 'employeeinformation'

    

def import_employee_information_to_SQL(source=None):

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
    
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before truncating')
    
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.employeeinformation'''))
    session.commit()
    session.close()
    
    
    # get count of table before insert --> should be 0
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before importing')
    print_count_results(engine=engine, schema='dbo', table=table, suffix_text='before merge proc')
    
    ei.to_sql('employeeinformation', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after importing')
          
    
    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.merge_employeeinformation()")
    else:
        cursor.execute("call dbo.merge_employeeinformation(%s)", (source,))
    connection.commit()
    connection.close()
    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine=engine, schema='dbo', table=table, suffix_text='after merge proc') 
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after merge proc')

       

def determine_terminated_employees(source=None):
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
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before truncating')
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.employeeinformation'''))
    session.commit()
    session.close()
    
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before importing')
    
    ei_terminated.to_sql('employeeinformation', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after importing')
    
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=True, suffix_text='before updating IS terminated')
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=False, suffix_text='before updating IS terminated')

    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.update_employeeinformation_isterminated()")
    else:
        cursor.execute("call dbo.update_employeeinformation_isterminated(%s)", (source,))
    connection.commit()
    connection.close()
    
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=True, suffix_text='after updating IS terminated')
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=False, suffix_text='after updating IS terminated')

    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after merge proc')
    
    
    
    
    print('RUNNING UPDATES FOR NON-TERMINATED EMPLOYEES')
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before truncating')
    
    Session = sessionmaker(bind=engine)
    session = Session()
    session.execute(text('''TRUNCATE TABLE live.employeeinformation'''))
    session.commit()
    session.close()
    
    print_count_results(engine=engine, schema='live', table=table, suffix_text='before importing')
    
    ei_notTerminated.to_sql('employeeinformation', engine, schema='live', if_exists='replace', index=False)
    # get count of table after insert
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after importing')
    
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=True, suffix_text='before updating IS terminated')
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=False, suffix_text='before updating IS terminated')

    connection = engine.raw_connection()
    cursor = connection.cursor()
    if source is None:
        cursor.execute("call dbo.update_employeeinformation_isnnotterminated()")
    else:
        cursor.execute("call dbo.update_employeeinformation_isnnotterminated(%s)", (source,))
    connection.commit()
    connection.close()
    
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=True, suffix_text='after updating IS terminated')
    print_count_results_where(engine=engine, schema='dbo', table=table, fieldname='terminated', equal_to_value=False, suffix_text='after updating IS terminated')

    
    # double check length of live table after merge proc --> should be 0
    print_count_results(engine=engine, schema='live', table=table, suffix_text='after merge proc')
    