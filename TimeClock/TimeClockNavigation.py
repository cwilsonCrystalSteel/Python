# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:36:05 2024

@author: CWilson
"""

import time
import datetime
import os
import glob
import datetime
from TimeClock_Credentials import returnTimeClockCredentials
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException


def printwait(string='', waittime=0):
    print(string)
    time.sleep(waittime)


def delete_range(web_element, x=20):
    web_element.click()
    for i in range(x):
        web_element.send_keys(Keys.BACKSPACE)
        web_element.send_keys(Keys.DELETE)


def enable_download(driver, download_folder):
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath':download_folder}}
    driver.execute("send_command", params)


def setting_chrome_options(headless):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
        chrome_options.add_argument('--no-sandbox')
    return chrome_options;


def newest_creation_time(download_folder):
    files = [download_folder + file for file in os.listdir(download_folder)]
    if len(files) == 0:
        newest_file_ctime = datetime.datetime.now() - datetime.timedelta(days=1)
    else:
        newest_file = max(files, key=os.path.getctime)
        newest_file_ctime = datetime.datetime.fromtimestamp(os.path.getctime(newest_file))
    return newest_file_ctime



def validateElement(driver, findElementArg, elemOutputName, checkPresence=True, checkClickable=True, timeoutLimit=10):
    if checkPresence:
        try:
            elem = WebDriverWait(driver, timeoutLimit).until(EC.presence_of_element_located(findElementArg))
            print(f'The element is present:   {elemOutputName}')
        except TimeoutException:
            print(f'Timeout Error for presence: {elemOutputName}')  
    
    if checkClickable:
        try:
            elem = WebDriverWait(driver, timeoutLimit).until(EC.element_to_be_clickable(findElementArg))
            print(f'The element is clickable: {elemOutputName}')
        except TimeoutException:
            print(f'Timeout Error for clickability: {elemOutputName}')  
            
    return elem



#%%
class TimeClockBase():
    def __init__(self, download_folder='c:\\users\\cwilson\\downloads\\', headless=True, fullscreen=False):
        
        self.creds = returnTimeClockCredentials()
        self.download_folder = download_folder
        self.headless = headless
        self.fullscreen = fullscreen
        
        if self.fullscreen and self.headless:
            Exception('Both options fullscreen & headless are enabled, only one can be enabled at a time')
          
        # init the service?
        self.service = Service()
        
        # set options for headless browser or not
        if self.headless:
            print('Browser will be HEADLESS')
            options = setting_chrome_options(True)
        else:
            options = setting_chrome_options(False)
            
        # init dirver with headless/options
        self.driver = webdriver.Chrome(service=self.service, options = options)
        print('Driver created...')
            
        # go fullscreen - when not headless!
        if self.fullscreen:  
            print('Browser will be FULLSCREEN')
            # full screen cus this website sucks
            self.driver.maximize_window()
            
        # create download folder if not exists
        
        if os.path.exists(download_folder):
            enable_download(self.driver, self.download_folder)
        else:
            os.makedirs(download_folder)
            enable_download(self.driver, self.download_folder)
   
        
        
    def kill(self):
        print('Attempting to quit the browser...')
        self.driver.quit()
        # self.driver.close()
        print('Quitting browser completed!')
        
    def splashPage(self):
        print('Navigating to splashpage!')
        # navigate to timeclock website
        self.driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
        print('Splashpage reached...')
        
            
    def startupBrowser(self):
        self.splashPage()
        
        self.usernameInput = validateElement(self.driver, (By.ID, 'LogOnUserId'), 'usernameInput', checkPresence=True, checkClickable=True)
        self.passwordInput = validateElement(self.driver, (By.ID, 'LogOnUserPassword'), 'passwordInput', checkPresence=True, checkClickable=True)
            
        
    def login(self):
        # try and remove anything in the text boxes
        delete_range(self.usernameInput)
        delete_range(self.passwordInput)
        
        # enter the username
        self.usernameInput.send_keys(self.creds['username'])
        print('Entered Username')
        
        
        # enter password
        self.passwordInput.send_keys(self.creds['password'])
        print('Entered Password')
        
        # Press 'enter' to login
        self.passwordInput.send_keys(Keys.RETURN)
        
        
    def tryLogin(self):
        
        try:
            self.login()
        except ElementNotInteractableException:
            print('Elements of login() are not Interactable at this time, retrying startup..')
            self.startupBrowser()
        
    def openTabularMenu(self):
        
        if not self.driver.current_url == 'https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509':
            Exception('Cannot begin to open Tabular Menu b/c the URL is not correct')
        
        self.tabularMenu = validateElement(self.driver, (By.CLASS_NAME, 'HeaderMenuIcon'), 'tabularMenu', checkPresence=True, checkClickable=True)
        
        # click the tabular Menu
        self.tabularMenu.click()
        print('Tabular Menu Clicked on...')
        
        
    def searchFromTabularMenu(self, searchText=''):
        # group hours:              searchText = 'Group Hours'
        # employee information      searchText = 'export'
        
        self.tabularMenuSearchBox = validateElement(self.driver, (By.CLASS_NAME, 'SearchInput'), 'tabularMenuSearchBox', checkPresence=True, checkClickable=True)
        
        
        self.tabularMenuSearchBox.send_keys(searchText)
        self.tabularMenuSearchBox.send_keys(Keys.RETURN)
        print(f'Searched "{searchText}"') 
        
    def clickTabularMenuSearchResults(self, findText=''):
        # group hours:              findText = 'Hours > Group Hours'
        # employee information:     findText = 'Tools > Export'
        
        self.clickTabularMenuSearchXPATH = f"//*[contains(text(), '{findText}')]"
        
        # print(f'Attempting to find by.XPATH: {self.clickTabularMenuSearchXPATH}')
        
        self.targetedItem = validateElement(self.driver, (By.XPATH, self.clickTabularMenuSearchXPATH), findText, checkPresence=True, checkClickable=True)
            
        # Click on the desired Group Hours or Employee Information
        self.targetedItem.click()
        print(f'Clicked "{findText}"')
        
        
    def employeeLocationFinale(self, exclude_terminated=False):
        # we want a time to base our download window from
        self.startTime = datetime.datetime.now()
        
        self.exportType = validateElement(self.driver, (By.ID, 'selExportType'), 'exportType', checkPresence=True, checkClickable=True)
        
        # creates a selectable(?) from the dropdown
        selectedExportType = Select(self.exportType)
        # this is much better approach then typing in the text and hoping for the best
        selectedExportType.select_by_visible_text('Employee Information')
        print('Found Export Type: Employee Information')
        
        # Navigate to Export Templates
        self.exportTemplatesDropdown = validateElement(self.driver, (By.XPATH, "//*[contains(text(), 'Export Templates')]"), 'exportTemplatesDropdown', checkPresence=True, checkClickable=True)
        self.exportTemplatesDropdown.click()
        print('Found Export Templates')
        
        # clicks on my custom report called "emplyee locations"
        self.employeeLocationsDropdownItem = validateElement(self.driver, (By.XPATH, "//*[contains(text(), 'employee location')]"), 'employeeLocationsDropdownItem', checkPresence=True, checkClickable=True)
        self.employeeLocationsDropdownItem.click()
        print('Found custom tempalte: employee locations')
        
        
        if exclude_terminated == False:
            
            
            self.employeeFiltersMenu = validateElement(self.driver, (By.XPATH, "//input[@value='Employee Filter']"), 'employeeFiltersMenu', checkPresence=True, checkClickable=True)
            self.employeeFiltersMenu.click()
            print('Found employee filter button')
            
            self.excludeSuspendedButton = validateElement(self.driver, (By.XPATH, "//input[@id='chkExcludeSuspended']"), 'excludeSuspendedButton', checkPresence=True, checkClickable=True)              
            # is_selected() = True if the box is checked to exclude suspended employees
            if self.excludeSuspendedButton.is_selected():
                self.excludeSuspendedButton.click()
                print('Un-checked the Exclude Suspended checkbox')
            
            
            self.excludeTerminatedButton = validateElement(self.driver, (By.XPATH, "//input[@id='chkExcludeTerminated']"), 'excludeTerminatedButton', checkPresence=True, checkClickable=True)               
            if self.excludeTerminatedButton.is_selected():
                self.excludeTerminatedButton.click()
                print('Un-checked the Exclude Terminated checkbox')
        
        
        
            self.submitFilterButton = validateElement(self.driver, (By.XPATH, "//input[@value='Filter']"), 'submitExclusionsFilterButton', checkPresence=True, checkClickable=True)               
            self.submitFilterButton.click()
            print('Filtered the suspended & terminated')
        
        
        self.generateButton = validateElement(self.driver, (By.XPATH, "//input[@value='Generate']"), 'generateButton', checkPresence=True, checkClickable=True)               
        self.generateButton.click()
        print('Found Generate')
        
        self.downloadButton = validateElement(self.driver, (By.XPATH, "//input[@value='Download']"), 'downloadButton', checkPresence=True, checkClickable=True)               
        self.downloadButton.click()
        print('Began File Download')       
        
        
    def retrieveDownloadedFile(self, waitTime, fileType, searchText):
        # Group Hours:              fileType = '*.html'
        #                           searchText = 'Hours'
        # Employee Information:     fileType = '*.csv'
        #                           searchText = 'Employee Information'
        
        print(f"Searching for {fileType} with text {searchText} in {self.download_folder}")
        endTime = time.time() + waitTime
        while True:
            
            self.listOfFileTypes = glob.glob(self.download_folder + fileType) # * means all if need specific format then *.csv
            # Create a list with only the states we want to look at
            self.foundFiles = [f for f in self.listOfFileTypes if searchText in f]
            # Get the most recent file for that state
            self.newestFile = max(self.foundFiles, key=os.path.getctime)
            # get the file creation time as a datetime
            self.newestFileTime = datetime.datetime.fromtimestamp(os.path.getctime(self.newestFile))
            
            if self.newestFileTime > self.startTime:
                print(f'File found: {self.newestFile}')
                return self.newestFile
            
            if time.time() > endTime:
                raise Exception('No reasonable download was found...')
            
        
            
            
        
        
x = TimeClockBase(headless=False)     
x.startupBrowser()
x.tryLogin()
x.openTabularMenu()
x.searchFromTabularMenu('export')
x.clickTabularMenuSearchResults('Tools > Export')
x.employeeLocationFinale()
filepath = x.retrieveDownloadedFile(10, '*.csv', 'Employee Information')

# x.searchFromTabularMenu('Group Hours')
# x.clickTabularMenuSearchResults('Hours > Group Hours')
# x.groupHoursFinale()
# filepath = x.retrieveDownloadedFile(10, '*.html', 'Hours')


  #%%      
x.kill()        

    
        
        
        
        
        
        
        
        
        