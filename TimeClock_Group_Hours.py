# -*- coding: utf-8 -*-
"""
Created on Tue Mar  9 15:36:31 2021

@author: CWilson
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
from TimeClock_scraping_functions import printwait, delete_range, enable_download, setting_chrome_options, newest_creation_time
import datetime
# import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from TimeClock_Credentials import returnTimeClockCredentials
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service


def download_group_hours(start_date, end_date, download_folder="C:\\Users\\cwilson\\Downloads\\"):
    
    print('Running Group Hours for: ' + start_date + ' to ' + end_date)
    
    # setup the driver
    service = Service()
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)
    
    # full screen cus this website sucks
    driver.maximize_window()
    
    # allow downloading
    enable_download(driver, download_folder)
        
    
    # navigate to timeclock website
    driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
    printwait('Loaded Website', 6)

    # Find the username field
    userid = driver.find_element(By.ID, 'LogOnUserId')
    # get the username & password dict
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
    searchbox.send_keys('Group Hours')
    searchbox.send_keys(Keys.RETURN)
    printwait('Searched "Group Hours"', 3) 
    
    
    # find group hours link
    grouphours = driver.find_element(By.XPATH, "//*[contains(text(), 'Hours > Group Hours')]")
    grouphours.click()
    printwait('Found Hours > Group Hours', 2)


    # find the Employee Filter button
    employeefilter = driver.find_element(By.XPATH, "//input[@value='Employee Filter']")
    employeefilter.click()   
    printwait('Clicked Employee Filter', 2)

    
    # find the suspended checkbox
    excludesuspended = driver.find_element(By.XPATH, "//input[@id='chkExcludeSuspended']")
    if excludesuspended.is_selected():
        excludesuspended.click()
        printwait('Clicked Suspended', 1)

    # excludeterminated.is_selected()
    # find the terminated checkbox
    excludeterminated = driver.find_element(By.XPATH, "//input[@id='chkExcludeTerminated']")
    if excludeterminated.is_selected():
        excludeterminated.click()
        printwait('Clicked Terminated', 1)

    
    # find the Filter button
    filterbutton = driver.find_element(By.XPATH, "//input[@value='Filter']")
    filterbutton.click()    
    printwait('Clicked Filter Button', 15)

   
    # Find the start date box
    startdate_box = driver.find_element(By.NAME, 'dpPeriodStart')
    # scroll to top of page b/c it didn't want to deal with the box without doing this 
    startdate_box.send_keys(Keys.CONTROL + Keys.HOME)
    # Clear the field
    delete_range(startdate_box)
    printwait('Deleted Start Date', 1)
    # Enter in date
    startdate_box.send_keys(start_date)
    printwait('Entered Start Date: ' + start_date, 3)
    # print('Entered Start Date')
    
    
    # Find the stop date box
    stopdate_box = driver.find_element(By.NAME, 'dpPeriodEnd')
    # Clear the field
    delete_range(stopdate_box)
    printwait('Deleted End Date', 1)
    # Enter in date
    stopdate_box.send_keys(end_date)
    printwait('Entered End Date: ' + end_date, 1)
    # print('Entered End Date')
    
 
        
    update_button = driver.find_element(By.XPATH, "//input[@value='Update']")
    update_button.click()
    printwait('Clicked update button', 5)
    
    
    try:
        noRecordsFoundText = driver.find_element(By.CLASS_NAME, 'NoDataListItem')
        print('Uh oh! we found text saying "No Records Found" for the search criteria... Adios!')
        # driver.quit()
        printwait('Closing browser', 2) 
        return None
    except:
        print('Good news, we did not find the text saying "No records found"')    
    
    
    def find_download():
    
        x = 0
        # max wait time is 30 * 10 seconds
        while x < 30:
            printwait('Trying to find downlaod button #1: ' + str(x), 10)
            x += 1
            try:
                # Find the Download Button
                download = driver.find_element(By.XPATH, "//*[contains(text(), 'Download')]")
                download.click()  
                printwait('Clicked Download #1', 1)
                break
            except:
                pass
    
    
        # Choose HTML download
        htmldownload = driver.find_elements(By.XPATH,  "//*[contains(text(), 'HTML')]")[0]
        htmldownload.click()
        printwait('Clicked HTML download type', 20)
        
    # do this loop to get the html download button
    find_download()
      

    # get the current timestamp to compare to
    now = datetime.datetime.now()
    # check when the newest file was downlaoded
    newest_file_ctime = newest_creation_time(download_folder)
    
    restart_message = "https://136509.tcplusondemand.com/app/manager/#/ManageHoursGroup - Form submission canceled because the form is not connected"
    
    
    x = 0
    # quit the loop if the newest file was created after 'now'
    while (now > newest_file_ctime) and (x < 30):
        
        printwait('Trying to find download button #2: ' + str(x), 10)
        
        browser_driver = driver.get_log('browser')
        print(browser_driver)
        
        if len(browser_driver):
        
            if browser_driver[0]['message'] == restart_message:
                 find_download()
        
        # try-catch for when it doesnt find the download button
        try:
            
            download2 = driver.find_element(By.XPATH, "//input[@value='Download']")
            time.sleep(5)
            download2.click()
            time.sleep(15)
            # check if a file was downloaded again
            # newest_file_ctime = newest_creation_time(download_folder)
            
        except:
            pass
        
        
        newest_file_ctime = newest_creation_time(download_folder)
        x += 1
    
    print('Began File Download')
    
    
    # time.sleep(30)
    print('Completed wait time to download')
    
    driver.quit()
    
    printwait('Closing browser', 2)
    
    return True

    








