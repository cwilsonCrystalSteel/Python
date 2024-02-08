# -*- coding: utf-8 -*-
"""
Created on Tue Mar  9 15:36:31 2021

@author: CWilson
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
# import chromedriver_autoinstaller
from TimeClock_scraping_functions import printwait, delete_range, enable_download, setting_chrome_options, newest_creation_time
from selenium.webdriver.common.by import By
from TimeClock_Credentials import returnTimeClockCredentials
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service



def download_most_current_employee_location_csv(download_folder="C:\\Users\\cwilson\\Downloads", exclude_terminated=True):
    # from datetime import date, timedelta
    # today = date.today()
    # the_day = today - timedelta(days = 0)
    # date = the_day.strftime("%m/%d/%Y")
    print('Running Employee Location Export')
    
    service = Service()
    driver = webdriver.Chrome(service=service, options = setting_chrome_options())
    
 
    enable_download(driver, download_folder)
    
    
    # navigate to timeclock website
    driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
    print('Loaded website')
    
    # Pause to fully load the page - had some issues without this command
    time.sleep(6)
    # Find the username field
    userid = driver.find_element(By.ID, 'LogOnUserId')
    timeclockCreds = returnTimeClockCredentials()
    # Submit username
    userid.send_keys(timeclockCreds['username'])
    printwait('Entered Username', 1)
    # Find the password field
    password = driver.find_element(By.ID, 'LogOnUserPassword')
    # Submit password
    password.send_keys(timeclockCreds['password'])
    printwait('Entered Password', 1)
    # Press 'enter' to login
    password.send_keys(Keys.RETURN)
    printwait('Submitted Login Detail', 5)
    
    tabularmenu = driver.find_element(By.CLASS_NAME, 'HeaderMenuIcon')
    printwait('Found Tabular menu', 3) 
    tabularmenu.click()
    printwait('Found search box', 3) 
    searchbox = driver.find_element(By.CLASS_NAME, 'SearchInput')
    searchbox.send_keys('export')
    searchbox.send_keys(Keys.RETURN)
    printwait('Searched "export"', 3) 
    
    
    time.sleep(5)
    
    toolsexport = driver.find_element(By.XPATH, "//*[contains(text(), 'Tools > Export')]")
    toolsexport.click()
    print('Found Tools > Export')
    
    time.sleep(2)
    
    
    # Find the Export Type drowpdown menu
    exporttype = driver.find_element(By.ID, 'selExportType')
    # type the name of the dropdown to select into the dropdown box
    exporttype.send_keys('Employee Information') # automatically selects once it finishes typing
    print('Found Export Type: Employee Information')
    time.sleep(2)
    
    # Navigate to Export Templates
    exporttemplates = driver.find_element(By.XPATH, "//*[contains(text(), 'Export Templates')]")
    exporttemplates.click()
    print('Found Export Templates')
    time.sleep(2)
    
    # clicks on my custom report called "emplyee locations"
    employeelocations = driver.find_element(By.XPATH, "//*[contains(text(), 'employee location')]")
    employeelocations.click()
    print('Found custom tempalte: employee locations')
    time.sleep(2)
    
    
    if exclude_terminated == False:
        employeefilter = driver.find_element(By.XPATH, "//input[@value='Employee Filter']")
        employeefilter.click()
        print('Found employee filter button')
        time.sleep(2)        
        
        
        excludesuspended = driver.find_element(By.XPATH, "//input[@id='chkExcludeSuspended']")
        print('Found the exclude suspended checkbox')
        # is_selected() = True if the box is checked to exclude suspended employees
        if excludesuspended.is_selected():
            excludesuspended.click()
            print('Un-checked the Exclude Suspended checkbox')
        time.sleep(2)
        
        
        excludeterminated = driver.find_element(By.XPATH, "//input[@id='chkExcludeTerminated']")
        print('Found the exclude terminated checkbox')
        if excludeterminated.is_selected():
            excludeterminated.click()
            print('Un-checked the Exclude Terminated checkbox')
        time.sleep(2)
    
        submitfilter = driver.find_element(By.XPATH, "//input[@value='Filter']")
        submitfilter.click()
        print('Filtered the suspended & terminated')
        time.sleep(2)
    
    
    generate = driver.find_element(By.XPATH, "//input[@value='Generate']")
    generate.click()
    print('Found Generate')
    time.sleep(10)
    
    # Navigate to and click the download button
    download2 = driver.find_element(By.XPATH, "//input[@value='Download']")
    download2.click()
    print('Began File Download')
    
    time.sleep(5)
    print('Completed wait time to download')
    
    driver.quit()
    print('Closing browser')
    
    
    time.sleep(2)
    








