# -*- coding: utf-8 -*-
"""
Created on Tue Mar  9 15:36:31 2021

@author: CWilson
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import glob
import os
from TimeClock_Credentials import returnTimeClockCredentials


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




def download_employee_hours(employee_names, start_date, end_date):
    # from datetime import date, timedelta
    # today = date.today()
    # the_day = today - timedelta(days = 0)
    # date = the_day.strftime("%m/%d/%Y")
    

    
    # Navigate to TimeClock & login

    # driver = webdriver.Chrome(executable_path='C:/Users/cwilson/Documents/Python/chromedriver.exe')

    # Start the browser
    driver = webdriver.Chrome(executable_path='C:/Users/cwilson/Documents/Python/chromedriver.exe', 
                              options = setting_chrome_options())
    enable_download(driver)
    
    
    # navigate to timeclock website
    driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
    print('Loaded website')
    print()
    
    # Pause to fully load the page - had some issues without this command
    time.sleep(6)
    # Find the username field
    userid = driver.find_element_by_id('LogOnUserId')
    timeclockCreds = returnTimeClockCredentials()
    # Submit username
    userid.send_keys(timeclockCreds['username'])
    print('Entered Username')
    # Find the password field
    password = driver.find_element_by_id('LogOnUserPassword')
    # Submit password
    password.send_keys(timeclockCreds['password'])
    print('Entered Password')
    # Press 'enter' to login
    password.send_keys(Keys.RETURN)
    
    time.sleep(5)
    
    
    
    searchbox = driver.find_element_by_xpath("//input[@name='searchInput']")
    searchbox.send_keys('Individual Hours')
    searchbox.send_keys(Keys.RETURN)
    print('Found Search Box - Searched for "Individual Hours"')
    
    time.sleep(5)
    
    individualhours = driver.find_element_by_xpath("//*[contains(text(), 'Hours > Individual Hours')]")
    individualhours.click()
    print('Found Hours > Individual Hours')
    
    time.sleep(2)    
    
    employee_html_list = []
    
    
    for employee_name in employee_names:
        print('Running Individual Employee Hours for: ' + str(employee_name) + ' - ' + start_date + ' to ' + end_date)
        # there are 2 items called 'searchInput'
        employeesearch = driver.find_elements_by_name('searchInput')
        # the second item is the search box we want
        for i in range(0,40):
            employeesearch[1].send_keys(Keys.BACKSPACE)
        
        employeesearch[1].send_keys(employee_name)    
        print('searched for: ' + employee_name)
        employeesearch[1].send_keys(Keys.RETURN)
        
        time.sleep(1)
        employeename = driver.find_element_by_xpath("//*[contains(text(), '" + employee_name +"')]")
        employeename.click()
        
        # let the employee page load
        time.sleep(4)
        
        # Find the start date box
        startdate_box = driver.find_element_by_name('dpPeriodStart')
        # Clear the field
        delete_range(startdate_box)
        time.sleep(0.5)
        # Enter in date
        startdate_box.send_keys(start_date)
        print('Entered Start Date')
        time.sleep(1)
        
        
        # Find the stop date box
        stopdate_box = driver.find_element_by_name('dpPeriodEnd')
        # Clear the field
        delete_range(stopdate_box)
        time.sleep(0.5)
        # Enter in date
        stopdate_box.send_keys(end_date)
        print('Entered End Date')
        time.sleep(1)
        
        # find the update button
        udpatebutton = driver.find_element_by_xpath("//input[@value='Update']")
        # click the update button
        udpatebutton.click()
        print('Clicked update button')
        #this can take a while if the date range is long
        time.sleep(5)
        
        
        # Find the Download Button
        download = driver.find_element_by_xpath("//*[contains(text(), 'Download')]")
        download.click()
        print('Clicked Download')    
        
        
    
        time.sleep(0.5)
        
        # Choose HTML download
        htmldownload = driver.find_elements_by_xpath("//*[contains(text(), 'HTML')]")[0]
        htmldownload.click()
        print('Clicked HTML')    
        
        
        time.sleep(5)
        
        # Navigate to and click the download button
        download2 = driver.find_element_by_xpath("//input[@value='Download']")
        download2.click()
        print('Began File Download')
        
        time.sleep(5)
        print('Completed wait time to download')
        
        
        list_of_htmls = glob.glob("C://Users//Cwilson//downloads//*.html") # * means all if need specific format then *.csv
        # Create a list with only the states we want to look at
        employee_hours_htmls = [f for f in list_of_htmls if "Hours" in f]
        # Get the most recent file for that state
        latest_employee_hours = max(employee_hours_htmls, key=os.path.getctime)
        print('Newest employee hours (' + employee_name +'): ', latest_employee_hours)  
        
        employee_html_list.append(latest_employee_hours)         
        
        
        
        
    
    driver.close()
    print('Closing browser')
    
    
    time.sleep(2)
    
    return employee_html_list








