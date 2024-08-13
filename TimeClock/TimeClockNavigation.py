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
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException


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



def validateElement(driver, findElementArg, elemOutputName, checkPresence=True, checkClickable=True, timeoutLimit=10, verbosity=0):
    if checkPresence:
        try:
            elem = WebDriverWait(driver, timeoutLimit).until(EC.presence_of_element_located(findElementArg))
            if verbosity >= 2:
                print(f'The element is present:   {elemOutputName}')
        except TimeoutException:
            print(f'Timeout Error for presence: {elemOutputName}')  
    
    if checkClickable:
        try:
            elem = WebDriverWait(driver, timeoutLimit).until(EC.element_to_be_clickable(findElementArg))
            if verbosity >= 2:
                print(f'The element is clickable: {elemOutputName}')
        except TimeoutException:
            print(f'Timeout Error for clickability: {elemOutputName}')  
            
    return elem


class NoRecordsFoundException(Exception):
    def __init__(self, message='The text "No Records Found" was observered'):
        # Call the base class constructor with the parameters it needs
        super(NoRecordsFoundException, self).__init__(message)
        

class DownloadButtonDisabledException(Exception):
    def __init__(self, message=''):
        # Call the base class constructor with the parameters it needs
        super(DownloadButtonDisabledException, self).__init__(message)


class WhileTimerTimeoutExcpetion(Exception):
    def __init__(self, message=''):
        # Call the base class constructor with the parameters it needs
        super(WhileTimerTimeoutExcpetion, self).__init__(message)


#%%
class TimeClockBase():
    def __init__(self, download_folder='c:\\users\\cwilson\\downloads\\', headless=False, fullscreen=False, offscreen=False):
        self.verbosity = 1
        '''
        verbosity = 0: only get messages about browser starting & download 
        verbosity = 1: get text on actions being clicked / strings entered / etc.
        verbosity = 2: get text on all element validation efforts plus verbosity = 1
        '''
        
        self.creds = returnTimeClockCredentials()
        self.download_folder = download_folder
        self.headless = headless
        self.fullscreen = fullscreen
        self.offscreen = offscreen
        
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
        
        if self.offscreen:
            self.moveOffscreen()
            
            
        # go fullscreen - when not headless!
        if self.fullscreen:  
            print('Browser will be FULLSCREEN')
            # full screen cus this website sucks
            self.maximizeWindow()
            
        
            
            
        # create download folder if not exists
        
        if os.path.exists(download_folder):
            enable_download(self.driver, self.download_folder)
        else:
            os.makedirs(download_folder)
            enable_download(self.driver, self.download_folder)
   
    def printverbosity(self, string):
        if self.verbosity >= 1:
            print(string)
            
    def moveOffscreen(self):
        print('moving offscreen')
        self.driver.set_window_position(-1920,0)
        
    def moveOnscreen(self):
        print('moving onscreen')
        self.driver.set_window_position(10,10)        
   
    def maximizeWindow(self):
        self.printverbosity('Maximizing window')
        self.driver.maximize_window()
    
    def zoom(self, percentage):
        self.percentage = percentage
        self.printverbosity(f'Browser will be at {self.percentage}% ZOOM')
        self.driver.execute_script(f"document.body.style.zoom='{self.percentage}%'")    
       
    def kill(self):
        print('Attempting to quit the browser...')
        self.driver.quit()
        # self.driver.close()
        print('Quitting browser completed!')
        
    def splashPage(self):
        self.printverbosity('Navigating to splashpage!')
        # navigate to timeclock website
        self.driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
        self.printverbosity('Splashpage reached...')
        
            
    def startupBrowser(self):
        self.splashPage()
        
        self.usernameInput = validateElement(self.driver, (By.ID, 'LogOnUserId'), 'usernameInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.passwordInput = validateElement(self.driver, (By.ID, 'LogOnUserPassword'), 'passwordInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
            
        
    def login(self):
        # try and remove anything in the text boxes
        delete_range(self.usernameInput)
        delete_range(self.passwordInput)
        
        # enter the username
        self.usernameInput.send_keys(self.creds['username'])
        self.printverbosity('Entered Username')
        
        
        # enter password
        self.passwordInput.send_keys(self.creds['password'])
        self.printverbosity('Entered Password')
        
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
        
        self.tabularMenu = validateElement(self.driver, (By.CLASS_NAME, 'HeaderMenuIcon'), 'tabularMenu', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        
        # click the tabular Menu
        self.tabularMenu.click()
        self.printverbosity('Tabular Menu Clicked on...')
        
        
    def searchFromTabularMenu(self, searchText=''):
        # group hours:              searchText = 'Group Hours'
        # employee information      searchText = 'export'
        
        self.tabularMenuSearchBox = validateElement(self.driver, (By.CLASS_NAME, 'SearchInput'), 'tabularMenuSearchBox', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        
        
        self.tabularMenuSearchBox.send_keys(searchText)
        self.tabularMenuSearchBox.send_keys(Keys.RETURN)
        self.printverbosity(f'Searched "{searchText}"') 
        
    def clickTabularMenuSearchResults(self, findText=''):
        # group hours:              findText = 'Hours > Group Hours'
        # employee information:     findText = 'Tools > Export'
        
        self.clickTabularMenuSearchXPATH = f"//*[contains(text(), '{findText}')]"
        
        # print(f'Attempting to find by.XPATH: {self.clickTabularMenuSearchXPATH}')
        
        self.targetedItem = validateElement(self.driver, (By.XPATH, self.clickTabularMenuSearchXPATH), findText, checkPresence=True, checkClickable=True, verbosity=self.verbosity)
            
        # Click on the desired Group Hours or Employee Information
        self.targetedItem.click()
        self.printverbosity(f'Clicked "{findText}"')
    
        
    
    def includeTerminatedSuspendedEmployees(self):
        self.employeeFiltersMenu = validateElement(self.driver, (By.XPATH, "//input[@value='Employee Filter']"), 'employeeFiltersMenu', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.employeeFiltersMenu.click()
        self.printverbosity('Found employee filter button')
        
        self.excludeSuspendedButton = validateElement(self.driver, (By.XPATH, "//input[@id='chkExcludeSuspended']"), 'excludeSuspendedButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)              
        # is_selected() = True if the box is checked to exclude suspended employees
        if self.excludeSuspendedButton.is_selected():
            self.excludeSuspendedButton.click()
            self.printverbosity('Un-checked the Exclude Suspended checkbox')
        
        
        self.excludeTerminatedButton = validateElement(self.driver, (By.XPATH, "//input[@id='chkExcludeTerminated']"), 'excludeTerminatedButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        if self.excludeTerminatedButton.is_selected():
            self.excludeTerminatedButton.click()
            self.printverbosity('Un-checked the Exclude Terminated checkbox')
    

        self.submitFilterButton = validateElement(self.driver, (By.XPATH, "//input[@value='Filter']"), 'submitExclusionsFilterButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        self.submitFilterButton.click()
        self.printverbosity('Filtered the suspended & terminated')
        
        
    def employeeLocationFinale(self, exclude_terminated=False):
        # we want a time to base our download window from
        self.startTime = datetime.datetime.now()
        
        self.exportType = validateElement(self.driver, (By.ID, 'selExportType'), 'exportType', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        
        # creates a selectable(?) from the dropdown
        selectedExportType = Select(self.exportType)
        # this is much better approach then typing in the text and hoping for the best
        selectedExportType.select_by_visible_text('Employee Information')
        self.printverbosity('Found Export Type: Employee Information')
        
        # Navigate to Export Templates
        self.exportTemplatesDropdown = validateElement(self.driver, (By.XPATH, "//*[contains(text(), 'Export Templates')]"), 'exportTemplatesDropdown', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.exportTemplatesDropdown.click()
        self.printverbosity('Found Export Templates')
        
        # clicks on my custom report called "emplyee locations"
        self.employeeLocationsDropdownItem = validateElement(self.driver, (By.XPATH, "//*[contains(text(), 'employee location')]"), 'employeeLocationsDropdownItem', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.employeeLocationsDropdownItem.click()
        self.printverbosity('Found custom tempalte: employee locations')
        
        # termianted & suspended employees are automatically excluded
        if exclude_terminated == False:
            self.includeTerminatedSuspendedEmployees()
            
           
        
        
        self.generateButton = validateElement(self.driver, (By.XPATH, "//input[@value='Generate']"), 'generateButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        try:
            self.generateButton.click()
        except ElementClickInterceptedException:
            self.generatedButton.send_keys(Keys.TAB)
            self.generateButton.click()
        self.printverbosity('Found Generate')
        
        self.downloadButton = validateElement(self.driver, (By.XPATH, "//input[@value='Download']"), 'downloadButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        self.downloadButton.click()
        self.printverbosity('Began File Download')       
        
        
    def retrieveDownloadedFile(self, waitTime, fileType, searchText):
        # Group Hours:              fileType = '*.html'
        #                           searchText = 'Hours'
        # Employee Information:     fileType = '*.csv'
        #                           searchText = 'Employee Information'
        
        self.printverbosity(f"Searching for {fileType} with text {searchText} in {self.download_folder}")
        endTime = time.time() + waitTime
        while True:
            try:
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
            except Exception as e:
                # mostly this will just be 'max() arg is an empty sequence' 
                # becuase the file will not be downloaded yet
                # print(e)
                pass
                
            if time.time() > endTime:
                # raise Exception('No reasonable download was found...')
                raise WhileTimerTimeoutExcpetion(f'No filetype {fileType} was found downloaded.')
    
                
    def waitForProcessingPopup(self):
        # need to find the processing popup
        self.processingPopup = validateElement(self.driver, (By.CLASS_NAME, 'ProgressIndicatorModal'), 'processingPopup', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)               
        # then need to wait til its gone
        
        # give it 15 seconds or until the popup box is not available, just rapid fire check
        endTime = time.time() + 15
        while True:
            try:
                self.processingPopupStillAvailable = self.driver.find_element(By.CLASS_NAME, 'ProgressIndicatorModal')
            except NoSuchElementException:
                self.printverbosity('The processingPopup box has disappeared!')
                break
            
            if time.time() > endTime:
                # raise Exception('The processingPopup box never went away!')
                raise WhileTimerTimeoutExcpetion('The processingPopup box never went away!')
                
    def groupHoursFinale(self, dateString, exclude_terminated=False):
        self.startTime = datetime.datetime.now()
        
        
        # termianted & suspended employees are automatically excluded
        if exclude_terminated == False:
            self.includeTerminatedSuspendedEmployees()
    
        self.waitForProcessingPopup()
        
        
        
        # Find the stop date box
        self.endDateInput = validateElement(self.driver, (By.NAME, 'dpPeriodEnd'), 'endDateInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        # find the start date box
        self.startDateInput = validateElement(self.driver, (By.NAME, 'dpPeriodStart'), 'startDateInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        
        
        
        # Clear the field
        delete_range(self.endDateInput)
        self.printverbosity('Deleted End Date')
        # Enter in date
        self.endDateInput.send_keys(dateString)
        self.printverbosity(f'Entered End Date: {dateString}')
        
        
        delete_range(self.startDateInput)
        self.printverbosity('Deleted Start Date')
        # Enter in date
        self.startDateInput = validateElement(self.driver, (By.NAME, 'dpPeriodStart'), 'startDateInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        self.startDateInput.send_keys(dateString)
        self.printverbosity(f'Entered Start Date: {dateString}')

        
     
        self.updateButton = validateElement(self.driver, (By.XPATH, "//input[@value='Update']"), 'updateButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        # self.updateButton = self.driver.find_element(By.XPATH, "//input[@value='Update']")
        self.updateButton.click()
        self.printverbosity('Clicked update button')
        
        self.waitForProcessingPopup()
        
        try:
            self.noRecordsFoundText = validateElement(self.driver, (By.CLASS_NAME, 'NoDataListItem'), 'noRecordsFoundText', checkPresence=True, checkClickable=False, verbosity=self.verbosity)               
            # if this works, then it will close the browser. If it does not, it will error and hit the exception
            # self.noRecordsFoundText = self.driver.find_element(By.CLASS_NAME, 'NoDataListItem')
            self.printverbosity('Uh oh! we found text saying "No Records Found" for the search criteria...')
            
        except:
            self.printverbosity('Good news, we did not find the text saying "No records found"')    
    
    
        try:
            self.noRecordsFoundText
        except AttributeError:
            self.noRecordsFoundTextExists = False
            # throw excpetion if we find the no records
        else:
            self.noRecordsFoundTextEsists = True
        
            # raise Exception('NoRecordsFoundException')
            raise NoRecordsFoundException()
            
    
        # self.zoom(50)
        
        # self.maximizeWindow()
    
        # time.sleep(1)
    
        ''' THIS IS WHERE I AM LEAVING OFF
        
        I was trying to do the Zoom out so that you could see the download button, but after 
        zooming out, it is not interactable ElementNotInteractableException for some reason
        
        '''
        
        # self.menuDownloadButton = validateElement(self.driver, (By.CLASS_NAME, 'Download'), 'menuDownloadButton', checkPresence=True, checkClickable=True)
        # trying to get it into view  
        self.driver.execute_script("window.scrollBy(10000,-1000);")
        self.menuDownloadButton = validateElement(self.driver, (By.CLASS_NAME, 'DownloadMenu'), 'menuDownloadButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.menuDownloadButtonDisabled = self.menuDownloadButton.get_attribute('disabled')
        if self.menuDownloadButtonDisabled is not None:
            # raise Exception('menuDownloadButtonDisabled')
            raise DownloadButtonDisabledException('menuDownloadButton')


        endTime = time.time() + 15
        while True:
            try:
                self.menuDownloadButton.click()
                self.printverbosity('We were able to press the menuDownloadButton')
                break
            except Exception as e:
                print(e)
                
            
            if time.time() > endTime:
                # raise Exception('We could not press the menuDownloadButton')
                raise WhileTimerTimeoutExcpetion('We could not press the menuDownloadButton')
            
        # self.menuDownloadButton.click()
        # self.menuDownloadButton.send_keys(Keys.RETURN)
        
        
        # self.processingPopup = validateElement(self.driver, (By.CLASS_NAME, 'ProgressIndicatorModal'), 'processingPopup', checkPresence=True, checkClickable=True, timeoutLimit=15)               
        self.htmlOption = self.driver.find_elements(By.XPATH,  "//*[contains(text(), 'HTML')]")[0]
        self.htmlOption.click()
        printwait('Clicked HTML download type', 0)
        
        
        
        self.processingPopup = validateElement(self.driver, (By.CLASS_NAME, 'ProgressIndicatorModal'), 'processingPopup', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)               
        # self.processingPopupDownloadButton = validateElement(self.driver, (By.CLASS_NAME, 'DownloadMenu'), 'processingPopupDownloadButton', checkPresence=True, checkClickable=True, timeoutLimit=15)    

        endTime = time.time() + 15
        while True:
            try:
                self.processingPopupDownloadButton = validateElement(self.driver, (By.XPATH, "//input[@value='Download']"), 'processingPopupDownloadButton', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)    
                self.processingPopupDownloadButtonDisabled = self.processingPopupDownloadButton.get_attribute('disabled')
                if self.processingPopupDownloadButtonDisabled is not None:
                    # raise Exception('processingPopupDownloadButtonDisabled')    
                    raise DownloadButtonDisabledException('processsingPopupDownloadButton')
                else:
                    self.processingPopupDownloadButton = validateElement(self.driver, (By.XPATH, "//input[@value='Download']"), 'processingPopupDownloadButton', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)    
                    self.processingPopupDownloadButton.click()
                    self.printverbosity('Processing Popup Download Button Clicked!')                    
                    break
                
            except Exception as e:
                print(e)
                
            
            if time.time() > endTime:
                # raise Exception('We could not press the processingPopupDownloadButton')
                raise WhileTimerTimeoutExcpetion('We could not press the processingPopupDownloadButton')
                
                
        # self.processingPopupDownloadButtonDisabled = self.processingPopupDownloadButton.get_attribute('disabled')
        # if self.processingPopupDownloadButton is not None:
        #     raise Exception('processingPopupDownloadButtonDisabled')
            
        # self.processingPopupDownloadButton.click()
        # print('Processing Popup Download Button Clicked!')
        
        
        # self.download2 = self.driver.find_element(By.XPATH, "//input[@value='Download']")
        
        
class TimeClockEZGroupHours(TimeClockBase):
    def __init__(self, date_str):
        
        
        self.date_str = date_str
        
    def get_filepath(self):
        self.tcb = TimeClockBase(download_folder="C:\\users\\cwilson\\downloads\\GroupHours\\", offscreen=True)
        self.tcb.verbosity=1
        self.tcb.startupBrowser()
        self.tcb.tryLogin()
        self.tcb.openTabularMenu()
        self.tcb.searchFromTabularMenu('Group Hours')
        self.tcb.clickTabularMenuSearchResults('Hours > Group Hours')
        try:
            self.tcb.groupHoursFinale(self.date_str)
            self.filepath = self.tcb.retrieveDownloadedFile(10, '*.html', 'Hours')
        except (DownloadButtonDisabledException, NoRecordsFoundException)  as e:
            print(f'Could not accomplish {self.date_str} because of {e}')
            self.filepath = e
        except WhileTimerTimeoutExcpetion as e:
            print(f'Could not accomplish {self.date_str} because of {e}')
            print('perhaps you should try again')
            self.filepath = None
        
        
        
        
        return self.filepath
    
    def kill(self):
        self.tcb.kill()
        
        
'''
x = TimeClockEZGroupHours('07/21/2024')
filepath = x.get_filepath()
x.kill()
'''        
        
        
        
'''        
x = TimeClockBase(offscreen=False)  
# x.maximizeWindow()
# x.moveOnscreen()
x.verbosity=2
x.startupBrowser()
x.tryLogin()
x.openTabularMenu()
# x.searchFromTabularMenu('export')
# x.clickTabularMenuSearchResults('Tools > Export')
# x.employeeLocationFinale()
# filepath = x.retrieveDownloadedFile(10, '*.csv', 'Employee Information')

x.searchFromTabularMenu('Group Hours')
x.clickTabularMenuSearchResults('Hours > Group Hours')
x.groupHoursFinale('07/21/2024')
filepath = x.retrieveDownloadedFile(10, '*.html', 'Hours')
print(filepath)


  #%%      
x.kill()        

'''
        
        
        
        
        
        
        
        
        