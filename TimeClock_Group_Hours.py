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

class download_group_hours():
    def __init__(self, start_date, end_date, download_folder="C:\\Users\\cwilson\\Downloads\\"):
# def download_group_hours(start_date, end_date, download_folder="C:\\Users\\cwilson\\Downloads\\"):
    
        self.start_date = start_date
        self.end_date = end_date
        self.download_folder = download_folder
        
        
    def startup(self):
        
        print('Running Group Hours for: ' + self.start_date + ' to ' + self.end_date)
        
        # setup the driver
        service = Service()
        options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(service=service, options=options)
        
        # full screen cus this website sucks
        self.driver.maximize_window()
        
        # allow downloading
        enable_download(self.driver, self.download_folder)
            
        
        # navigate to timeclock website
        self.driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
        printwait('Loaded Website', 6)
        
        # Find the username field
        userid = self.driver.find_element(By.ID, 'LogOnUserId')
        # get the username & password dict
        timeclockCreds = returnTimeClockCredentials()
        # Submit username
        userid.send_keys(timeclockCreds['username'])
        printwait('Entered Username', 1)
        # Find the password field
        password = self.driver.find_element(By.ID, 'LogOnUserPassword')
        # Submit password
        password.send_keys(timeclockCreds['password'])
        printwait('Entered Password', 1)
        # Press 'enter' to login
        password.send_keys(Keys.RETURN)
        printwait('Submitted Login Detail', 5)
        
        # we made it ! say so
        return True
        
        
    def navigate(self):
        
        
        self.tabularmenu = self.driver.find_element(By.CLASS_NAME, 'HeaderMenuIcon')
        printwait('Found Tabular menu', 3) 
        self.tabularmenu.click()
        
        
        printwait('Found search box', 3) 
        self.searchbox = self.driver.find_element(By.CLASS_NAME, 'SearchInput')
        self.searchbox.send_keys('Group Hours')
        self.searchbox.send_keys(Keys.RETURN)
        printwait('Searched "Group Hours"', 3) 
        
        
        # find group hours link
        self.grouphours = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Hours > Group Hours')]")
        self.grouphours.click()
        printwait('Found Hours > Group Hours', 2)
    
    
        # find the Employee Filter button
        self.employeefilter = self.driver.find_element(By.XPATH, "//input[@value='Employee Filter']")
        self.employeefilter.click()   
        printwait('Clicked Employee Filter', 2)
    
        
        # find the suspended checkbox
        self.excludesuspended = self.driver.find_element(By.XPATH, "//input[@id='chkExcludeSuspended']")
        if self.excludesuspended.is_selected():
            self.excludesuspended.click()
            printwait('Clicked Suspended', 1)
    
        # excludeterminated.is_selected()
        # find the terminated checkbox
        self.excludeterminated = self.driver.find_element(By.XPATH, "//input[@id='chkExcludeTerminated']")
        if self.excludeterminated.is_selected():
            self.excludeterminated.click()
            printwait('Clicked Terminated', 1)
    
        
        # find the Filter button
        self. filterbutton = self.driver.find_element(By.XPATH, "//input[@value='Filter']")
        self.filterbutton.click()    
        printwait('Clicked Filter Button', 15)
    
       
        # Find the start date box
        self.startdate_box = self.driver.find_element(By.NAME, 'dpPeriodStart')
        # scroll to top of page b/c it didn't want to deal with the box without doing this 
        self.startdate_box.send_keys(Keys.CONTROL + Keys.HOME)
        # Clear the field
        delete_range(self.startdate_box)
        printwait('Deleted Start Date', 1)
        # Enter in date
        self.startdate_box.send_keys(self.start_date)
        printwait('Entered Start Date: ' + self.start_date, 3)
        # print('Entered Start Date')
        
        
        # Find the stop date box
        self.stopdate_box = self.driver.find_element(By.NAME, 'dpPeriodEnd')
        # Clear the field
        delete_range(self.stopdate_box)
        printwait('Deleted End Date', 1)
        # Enter in date
        self.stopdate_box.send_keys(self.end_date)
        printwait('Entered End Date: ' + self.end_date, 1)
        # print('Entered End Date')
        
     
            
        self.update_button = self.driver.find_element(By.XPATH, "//input[@value='Update']")
        self.update_button.click()
        printwait('Clicked update button', 5)
        
        # try and see if we can find the text 'No Records Found'
        try:
            # if this works, then it will close the browser. If it does not, it will error and hit the exception
            self.noRecordsFoundText = self.driver.find_element(By.CLASS_NAME, 'NoDataListItem')
            print('Uh oh! we found text saying "No Records Found" for the search criteria... Adios!')
            # self.driver.quit()
            printwait('Closing browser', 2) 
            self.driver.quit()
            return False
        except:
            print('Good news, we did not find the text saying "No records found"')    
    
        # try and see if the download button is available
        # and see if we can see if it has the 'disabled' attriute
        try:
            self.download1a = self.driver.find_element(By.CLASS_NAME, "DownloadMenu")
            self.download1aDisabled = self.download1a.get_attribute('disabled')
            print(f"Attribute of self.download1a (in navigate): {self.download1aDisabled}")
            # I think we only see this value if it is true
            if self.download1aDisabled is not None:
                print('we found that the attribute of the download1 box was DISABLED... Adios!')
                self.driver.quit()
                return False
        except:
            print('Good news, we found the download button but did not see the disabled tag')
            
            
        # just give somethign if we made it this far
        return True
        
        
    def find_download(self):
    
        x = 0
        # max wait time is 30 * 10 seconds
        while x < 30:
            printwait('Trying to find downlaod button #1: ' + str(x), 10)
            x += 1
            try:
                # Find the Download Button
                # download = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Download')]")
                
                self.download1b = self.driver.find_element(By.CLASS_NAME, "DownloadMenu")
                self.download1bDisabled = self.download1b.get_attribute('disabled')
                print(f"Attribute of self.download1b (in find_download): {self.download1bDisabled}")
                
                if self.download1bDisabled is not None:
                    self.driver.quit()
                    return False
                
                
                self.download1b.click()  
                printwait('Clicked Download #1', 1)
                break
            except:
                pass
    
    
        # Choose HTML download
        self.htmldownload = self.driver.find_elements(By.XPATH,  "//*[contains(text(), 'HTML')]")[0]
        self.htmldownload.click()
        printwait('Clicked HTML download type', 20)
        
        # pass True if we reach this point
        return True
        
        
        
    def downloader(self):
    
        self.startupResults = self.startup()
        self.navigateResults = self.navigate()
        if not self.navigateResults:
            return False
      
        # do this loop to get the html download button
        self.findDownloadResult = self.find_download()
        if not self.findDownloadResult:
            return False
    
        # get the current timestamp to compare to
        now = datetime.datetime.now()
        # check when the newest file was downlaoded
        newest_file_ctime = newest_creation_time(self.download_folder)
        
        restart_message = "https://136509.tcplusondemand.com/app/manager/#/ManageHoursGroup - Form submission canceled because the form is not connected"
        
        
        x = 0
        # quit the loop if the newest file was created after 'now'
        while (now > newest_file_ctime) and (x < 30):
            
            printwait('Trying to find download button #2: ' + str(x), 10)
            
            browser_driver = self.driver.get_log('browser')
            print(browser_driver)
            
            if len(browser_driver):
            
                if browser_driver[0]['message'] == restart_message:
                     self.find_download()
            
            # try-catch for when it doesnt find the download button
            try:
                
                self.download2 = self.driver.find_element(By.XPATH, "//input[@value='Download']")
                time.sleep(5)
                self.download2.click()
                time.sleep(15)
                # check if a file was downloaded again
                # newest_file_ctime = newest_creation_time(download_folder)
                
            except:
                pass
            
            
            newest_file_ctime = newest_creation_time(self.download_folder)
            x += 1
        
        print('Began File Download')
        
        
        # time.sleep(30)
        print('Completed wait time to download')
        
        self.driver.quit()
        
        printwait('Closing browser', 2)
        
        return True
    
        








