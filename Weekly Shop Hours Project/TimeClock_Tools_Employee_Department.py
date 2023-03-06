# -*- coding: utf-8 -*-
"""
Created on Tue Mar  9 15:36:31 2021

@author: CWilson
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time


def delete_range(web_element, x=20):
    web_element.click()
    for i in range(x):
        web_element.send_keys(Keys.BACKSPACE)
        web_element.send_keys(Keys.DELETE)



def enable_download(driver):
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "C:\\Users\\cwilson\\Downloads"}}
    driver.execute("send_command", params)

def setting_chrome_options():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument('--no-sandbox')
    return chrome_options;




def download_most_current_employee_department_csv():
    # from datetime import date, timedelta
    # today = date.today()
    # the_day = today - timedelta(days = 0)
    # date = the_day.strftime("%m/%d/%Y")
    print('Running Employee Department Export')
    
    
    
    # Navigate to TimeClock & login

    # driver = webdriver.Chrome(executable_path='C:/Users/cwilson/Documents/Python/chromedriver.exe')

    # Start the browser
    driver = webdriver.Chrome(executable_path='C:/Users/cwilson/Documents/Python/chromedriver.exe', 
                              options = setting_chrome_options())
    enable_download(driver)

    
    # navigate to timeclock website
    driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
    print('Loaded website')
    
    # Pause to fully load the page - had some issues without this command
    time.sleep(6)
    # Find the username field
    userid = driver.find_element_by_id('LogOnUserId')
    # Submit username
    userid.send_keys('codyw')
    print('Entered Username')
    # Find the password field
    password = driver.find_element_by_id('LogOnUserPassword')
    # Submit password
    password.send_keys('CW8036V')
    print('Entered Password')
    # Press 'enter' to login
    password.send_keys(Keys.RETURN)
    
    time.sleep(5)
    
    
    
    searchbox = driver.find_element_by_xpath("//input[@name='searchInput']")
    searchbox.send_keys('export')
    searchbox.send_keys(Keys.RETURN)
    print('Found Search Box - Searched for "export"')
    
    time.sleep(5)
    
    toolsexport = driver.find_element_by_xpath("//*[contains(text(), 'Tools > Export')]")
    toolsexport.click()
    print('Found Tools > Export')
    
    time.sleep(2)
    
    # Find the Export Type drowpdown menu
    exporttype = driver.find_element_by_id('selExportType')
    # type the name of the dropdown to select into the dropdown box
    exporttype.send_keys('Employee Information') # automatically selects once it finishes typing
    print('Found Export Type: Employee Information')
    time.sleep(2)
    
    # Navigate to Export Templates
    exporttemplates = driver.find_element_by_xpath("//*[contains(text(), 'Export Templates')]")
    exporttemplates.click()
    print('Found Export Templates')
    time.sleep(2)
    
    # clicks on my custom report called "emplyee locations"
    employeelocations = driver.find_element_by_xpath("//*[contains(text(), 'Employee Departments')]")
    employeelocations.click()
    print('Found Custom Tempalte: Employee Departments')
    time.sleep(2)
    
    generate = driver.find_element_by_xpath("//input[@value='Generate']")
    generate.click()
    print('Found Generate')
    
    time.sleep(10)
    
    # Navigate to and click the download button
    download2 = driver.find_element_by_xpath("//input[@value='Download']")
    download2.click()
    print('Began File Download')
    
    time.sleep(5)
    print('Completed wait time to download')
    
    driver.close()
    print('Closing browser')
    
    
    time.sleep(2)
    








