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




def get_most_current_html(state='TN', start_date="03/06/2021", end_date="03/06/2021"):
    # from datetime import date, timedelta
    # today = date.today()
    # the_day = today - timedelta(days = 0)
    # date = the_day.strftime("%m/%d/%Y")
    print('Running TimeClock for: ' + state + " from " + start_date + " to " + end_date)
    
    
    
    # Navigate to TimeClock & login

    # driver = webdriver.Chrome(executable_path='C:/Users/cwilson/Documents/Python/chromedriver.exe')

    # Start the browser
    driver = webdriver.Chrome(executable_path='C:/Users/cwilson/Documents/Python/chromedriver.exe', 
                              options = setting_chrome_options())
    enable_download(driver)
    
    # driver.implicitly_wait(5)

    # navigate to timeclock website
    driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
    print('Loaded website')
    
    # Pause to fully load the page - had some issues without this command
    time.sleep(4)
    # Find the username field
    userid = driver.find_element_by_id('LogOnUserId')
    # Submit username
    userid.send_keys('jturner')
    print('Entered Username')
    # Find the password field
    password = driver.find_element_by_id('LogOnUserPassword')
    # Submit password
    password.send_keys('Justin32')
    print('Entered Password')
    # Press 'enter' to login
    password.send_keys(Keys.RETURN)
    
    # Should now be into TimeClock, at the Reports Page
    ''' I really need to put a line of code in to find the REPORTS button before continuing (3/30/2021) '''
    
    
    time.sleep(2)
    
    jobcodedetail = []

    while len(jobcodedetail) == 0:
        
        #refresh the page
        driver.refresh()
        
        # find jobcode
        jobcode = driver.find_elements_by_xpath("//*[contains(text(), 'Job Code')]")
        
        # trys again to find jobcode if it failed
        if len(jobcode) == 2:
            jobcode[1].click()
            time.sleep(1)
            
        else:
            count = 0
            while len(jobcode) < 2:
                time.sleep(1)
                # for some reason it does not like to find jobcode
                # so it will run and run and run bc of this loop 
                jobcode = driver.find_elements_by_xpath("//*[contains(text(), 'Job Code')]")
                print('\t Attempting to find jobcode')
                count += 1
                if count > 10:
                    # print('\t\t RESTARTING TIMECLOCK')
                    # driver.close()
                    # get_most_current_html(state, start_date, end_date)

                    # creating the [None, None] is my way to quit the program and throw and error down the line
                    jobcode = [None, None]
                
            jobcode[1].click()
        print('Found JobCode')
        
        time.sleep(2)
        # looks for the correct job code summary
        state_detail = "//*[contains(text(), '" + state + " Job Code Detail *')]"
        jobcodedetail = driver.find_elements_by_xpath(state_detail)
        print('Refreshed b/c jobcodedetail not found')
        # refreshes if it doesnt find the jobcodedetail
        

        
    print('Found jobcodedetail')

    jobcodedetail[0].click()
    
    
    time.sleep(3)
    
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
    
    # Filter to remove employees 2001,2015,2029 ??????
    
    
    # Find the Download Button
    download = driver.find_element_by_xpath("//*[contains(text(), 'Download')]")
    download.click()
    print('Clicked Download')
    
    time.sleep(0.5)
    
    # Choose HTML download
    htmldownload = driver.find_elements_by_xpath("//*[contains(text(), 'HTML')]")[0]
    htmldownload.click()
    print('Clicked HTML')
    
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
    








