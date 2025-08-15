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
from TimeClock.TimeClock_Credentials import returnTimeClockCredentials
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException
import signal
from psutil import Process, NoSuchProcess
from pathlib import Path


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
    if not isinstance(download_folder, str):
        download_folder = str(download_folder)
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath':download_folder}}
    driver.execute("send_command", params)


def setting_chrome_options(headless):
    chrome_options = Options()
    if headless:
        # chrome_options.add_argument("--headless")
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument("--window-size=1920,1080")  # Forces large viewport
        chrome_options.add_argument("--headless=new")  # Chrome 109+ headless mode
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

def safe_send_keys(element, text, driver, retries=3):
    for attempt in range(retries):
        try:
            element.click()  # optional: focus before typing
            element.clear()
            element.send_keys(text)
            return
        except ElementClickInterceptedException:
            print(f"Tooltip in the way, attempt {attempt+1}/{retries}")
            WebDriverWait(driver, 3).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "ui-tooltip-content"))
            )
    raise Exception("Could not send keys after retries")


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
    def __init__(self, download_folder=None, headless=False, fullscreen=False, offscreen=False):
        self.verbosity = 1
        '''
        verbosity = 0: only get messages about browser starting & download 
        verbosity = 1: get text on actions being clicked / strings entered / etc.
        verbosity = 2: get text on all element validation efforts plus verbosity = 1
        '''
        
        self.creds = returnTimeClockCredentials()
        if download_folder is None:
            download_folder = Path.home() / 'Downloads'
        self.download_folder = download_folder
        self.headless = headless
        self.fullscreen = fullscreen
        self.offscreen = offscreen
        
        if self.fullscreen and self.headless:
            Exception("Both options 'fullscreen' & 'headless' are enabled, only one can be enabled at a time")
          
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
            try:
                os.makedirs(download_folder)
            except PermissionError as e:
                print(f'There was an error with permissions of filepath:\n{e}')
                self.download_folder = Path.home() / 'Downloads' / 'PermissionErrorOverflowDownloads'
                os.makedirs(download_folder)
                
            enable_download(self.driver, self.download_folder)
            
            
        self.pid = self.driver.service.process.pid
        
        self.screenshotDirectory = None
        
    def enableScreenshots(self, directory):
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        self.screenshotDirectory = directory
        
        self.printverbosity(f"Enabling Screenshots to: {self.screenshotDirectory}")
        
            
    def take_screenshot(self, name):
        if isinstance(self.screenshotDirectory, str):
            self.screenshotDirectory = Path(self.screenshotDirectory)
            
        if not self.screenshotDirectory is None:
            
            if not '.png' in name:
                name = name + '.png'
                
            outfile = self.screenshotDirectory / name
            self.printverbosity(f"Screenshot taken: {outfile}")
                
            self.driver.get_screenshot_as_file(outfile)
            
            
   
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
        
        
        # try:
        #     os.kill(int(self.pid), signal.SIGTERM)
        #     print(f"Killed chrome using process: {int(self.pid)}")
        # except ProcessLookupError as ex:
        #     print(ex)
        #     pass
    
    
        
        self.p = Process(self.pid)
        print(self.p)
        print(f"Killing PID {self.pid}")
        
        self.driver.close()
        print('Attempting to quit the browser...')
        self.driver.quit()
        
        # self.driver.close()
        print('Quitting browser completed!')
        try:
            self.p.terminate()
        except NoSuchProcess:
            print(f'Could not kill PID {self.pid} b/c it is already gone')
        
        
        
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
        self.employeeLocationsDropdownItem = validateElement(self.driver, (By.XPATH, "//*[contains(text(), 'employee locations')]"), 'employeeLocationsDropdownItem', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
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
                self.listOfFileTypes = glob.glob(os.path.join(self.download_folder, fileType)) # * means all if need specific format then *.csv
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
                raise WhileTimerTimeoutExcpetion(f'No filetype {fileType} was found downloaded.')
    
                
    # def waitForProcessingPopup(self):
    #     # need to find the processing popup
    #     self.processingPopup = validateElement(self.driver, (By.CLASS_NAME, 'ProgressIndicatorModal'), 'processingPopup', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)               
    #     # then need to wait til its gone
        
    #     # give it 15 seconds or until the popup box is not available, just rapid fire check
    #     endTime = time.time() + 15
    #     while True:
    #         try:
    #             self.processingPopupStillAvailable = self.driver.find_element(By.CLASS_NAME, 'ProgressIndicatorModal')
    #         except NoSuchElementException:
    #             self.printverbosity('The processingPopup box has disappeared!')
    #             break
            
    #         if time.time() > endTime:
    #             raise WhileTimerTimeoutExcpetion('The processingPopup box never went away!')
                
                
    def waitForProcessingPopup(self):
        popup_locator = (By.CLASS_NAME, 'ProgressIndicatorModal')
    
        # Step 1: Try to detect the popup (short timeout just to see if it appears)
        try:
            self.processingPopup = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located(popup_locator)
            )
            self.printverbosity('Processing popup detected!')
        except TimeoutException:
            self.printverbosity('No processing popup appeared.')
            return False  # never appeared
    
        # Step 2: If it appeared, wait until it's gone
        endTime = time.time() + 15
        while True:
            try:
                self.driver.find_element(*popup_locator)
            except NoSuchElementException:
                self.printverbosity('The processing popup box has disappeared!')
                return True  # appeared and then disappeared
    
            if time.time() > endTime:
                raise WhileTimerTimeoutExcpetion('The processing popup never went away!')     
                

                
    def groupHoursFinale(self, dateString, exclude_terminated=False):
        self.startTime = datetime.datetime.now()
        
        
        # termianted & suspended employees are automatically excluded
        if exclude_terminated == False:
            self.includeTerminatedSuspendedEmployees()
    
    
        self.take_screenshot('1')
    
        self.waitForProcessingPopup()
        
        self.take_screenshot('2')
        
        
        # Find the stop date box
        self.endDateInput = validateElement(self.driver, (By.NAME, 'dpPeriodEnd'), 'endDateInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        # find the start date box
        self.startDateInput = validateElement(self.driver, (By.NAME, 'dpPeriodStart'), 'startDateInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        
        self.take_screenshot('3')
        
        safe_send_keys(self.endDateInput, dateString, self.driver)
        
        # Clear the field
        delete_range(self.endDateInput)
        self.printverbosity('Deleted End Date')
        # Enter in date
        self.endDateInput.send_keys(dateString)
        self.printverbosity(f'Entered End Date: {dateString}')
        
        self.take_screenshot('4')
        
        
        safe_send_keys(self.startDateInput, dateString, self.driver)
        
        delete_range(self.startDateInput)
        self.printverbosity('Deleted Start Date')
        # Enter in date
        self.startDateInput = validateElement(self.driver, (By.NAME, 'dpPeriodStart'), 'startDateInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        self.startDateInput.send_keys(dateString)
        self.printverbosity(f'Entered Start Date: {dateString}')
        
        self.take_screenshot('5')

        
     
        self.updateButton = validateElement(self.driver, (By.XPATH, "//input[@value='Update']"), 'updateButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)               
        # self.updateButton = self.driver.find_element(By.XPATH, "//input[@value='Update']")
        self.updateButton.click()
        self.printverbosity('Clicked update button')
        
        self.take_screenshot('6')
        
        self.waitForProcessingPopup()
        
        try:
            self.noRecordsFoundText = validateElement(self.driver, (By.CLASS_NAME, 'NoDataListItem'), 'noRecordsFoundText', checkPresence=True, checkClickable=False, verbosity=self.verbosity)               
            # if this works, then it will close the browser. If it does not, it will error and hit the exception
            # self.noRecordsFoundText = self.driver.find_element(By.CLASS_NAME, 'NoDataListItem')
            self.printverbosity('Uh oh! we found text saying "No Records Found" for the search criteria...')
            
            self.take_screenshot('7a')
            
        except:
            self.printverbosity('Good news, we did not find the text saying "No records found"')  
            
            self.take_screenshot('7b')
    
    
        try:
            self.noRecordsFoundText
        except AttributeError:
            self.noRecordsFoundTextExists = False
            # throw excpetion if we find the no records
        else:
            self.noRecordsFoundTextEsists = True
        
            raise NoRecordsFoundException()
            
    
        # self.zoom(50)
        
        # self.maximizeWindow()
    
        # time.sleep(1)
    
        ''' THIS IS WHERE I AM LEAVING OFF
        
        I was trying to do the Zoom out so that you could see the download button, but after 
        zooming out, it is not interactable ElementNotInteractableException for some reason
        
        '''
        self.take_screenshot('8')
        
        # self.menuDownloadButton = validateElement(self.driver, (By.CLASS_NAME, 'Download'), 'menuDownloadButton', checkPresence=True, checkClickable=True)
        # trying to get it into view  
        self.driver.execute_script("window.scrollBy(10000,-1000);")
        self.take_screenshot('9')
        self.menuDownloadButton = validateElement(self.driver, (By.CLASS_NAME, 'DownloadMenu'), 'menuDownloadButton', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.menuDownloadButtonDisabled = self.menuDownloadButton.get_attribute('disabled')
        if self.menuDownloadButtonDisabled is not None:
            raise DownloadButtonDisabledException('menuDownloadButton')

        # this is to try and wait until the download button is available 
        endTime = time.time() + 15
        while True:
            try:
                self.menuDownloadButton.click()
                self.printverbosity('We were able to press the menuDownloadButton')
                self.take_screenshot('10')
                break
            # except ElementClickInterceptedException as e:
            #     print(e)
                # self.maximizeWindow()
                # self.driver.set_window_size(1920, 1080)
            except Exception as e:
                print(e)
                
            
            if time.time() > endTime:
                raise WhileTimerTimeoutExcpetion('We could not press the menuDownloadButton')
            

        
        self.take_screenshot('11')
        
        # self.processingPopup = validateElement(self.driver, (By.CLASS_NAME, 'ProgressIndicatorModal'), 'processingPopup', checkPresence=True, checkClickable=True, timeoutLimit=15)               
        self.htmlOption = self.driver.find_elements(By.XPATH,  "//*[contains(text(), 'HTML')]")[0]
        self.htmlOption.click()
        printwait('Clicked HTML download type', 0)
        
        self.take_screenshot('12')
        
        self.processingPopup = validateElement(self.driver, (By.CLASS_NAME, 'ProgressIndicatorModal'), 'processingPopup', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)               
        # self.processingPopupDownloadButton = validateElement(self.driver, (By.CLASS_NAME, 'DownloadMenu'), 'processingPopupDownloadButton', checkPresence=True, checkClickable=True, timeoutLimit=15)    

        self.take_screenshot('13')

        endTime = time.time() + 15
        while True:
            try:
                self.processingPopupDownloadButton = validateElement(self.driver, (By.XPATH, "//input[@value='Download']"), 'processingPopupDownloadButton', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)    
                self.processingPopupDownloadButtonDisabled = self.processingPopupDownloadButton.get_attribute('disabled')
                if self.processingPopupDownloadButtonDisabled is not None: 
                    raise DownloadButtonDisabledException('processsingPopupDownloadButton')
                else:
                    self.take_screenshot('14')
                    self.processingPopupDownloadButton = validateElement(self.driver, (By.XPATH, "//input[@value='Download']"), 'processingPopupDownloadButton', checkPresence=True, checkClickable=True, timeoutLimit=15, verbosity=self.verbosity)    
                    self.processingPopupDownloadButton.click()
                    self.printverbosity('Processing Popup Download Button Clicked!')    
                    self.take_screenshot('15')
                    break
                
            except Exception as e:
                print(e)
                
            
            if time.time() > endTime:
                raise WhileTimerTimeoutExcpetion('We could not press the processingPopupDownloadButton')
                
       
                
        # self.processingPopupDownloadButtonDisabled = self.processingPopupDownloadButton.get_attribute('disabled')
        # if self.processingPopupDownloadButton is not None:
        #     raise Exception('processingPopupDownloadButtonDisabled')
            
        # self.processingPopupDownloadButton.click()
        # print('Processing Popup Download Button Clicked!')
        
        
        # self.download2 = self.driver.find_element(By.XPATH, "//input[@value='Download']")
        
        
class TimeClockEZGroupHours(TimeClockBase):
    def __init__(self, date_str, offscreen=True, headless=False):
        
        
        self.date_str = date_str
        self.offscreen = offscreen
        self.headless = headless
        self.debug = False
        
        # self.download_folder = "C:\\users\\cwilson\\downloads\\GroupHours\\"
        self.download_folder = Path.home() / 'Downloads' / 'GroupHours'
        
    def get_filepath(self):
        if self.offscreen:
            self.tcb = TimeClockBase(download_folder=self.download_folder, offscreen=self.offscreen, fullscreen=True)
            
        
        elif self.headless:
            self.tcb = TimeClockBase(download_folder=self.download_folder, headless=True)
        
        self.tcb.verbosity=1
        
        if self.debug:
            self.tcb.verbosity = 2
            self.tcb.screenshotDirectory = 'C:\\Users\\Netadmin\\Downloads\\New folder'
        
        self.tcb.startupBrowser()
        self.tcb.tryLogin()
        self.tcb.openTabularMenu()
        self.tcb.searchFromTabularMenu('Group Hours')
        self.tcb.clickTabularMenuSearchResults('Hours > Group Hours')
        try:
            self.tcb.groupHoursFinale(self.date_str)
            self.filepath = self.tcb.retrieveDownloadedFile(10, '*.html', 'Hours')
            # self.filepath = self.tcb.filepath
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
x = TimeClockEZGroupHours('03/06/2025')
x = TimeClockEZGroupHours('03/06/2025', headless=True, offscreen=False)
x.debug = True

filepath = x.get_filepath()
x.kill()
'''        
        
        
        
'''        
self = TimeClockBase(offscreen=False)
self.screenshotDirectory = 'C:\\Users\\Netadmin\\Downloads\\New folder'
# self.maximizeWindow()
# self.moveOnscreen()
self.verbosity=2
self.startupBrowser()
self.tryLogin()
self.openTabularMenu()
# self.searchFromTabularMenu('export')
# self.clickTabularMenuSearchResults('Tools > Export')
# self.employeeLocationFinale()
# filepath = self.retrieveDownloadedFile(10, '*.csv', 'Employee Information')

self.searchFromTabularMenu('Group Hours')
self.clickTabularMenuSearchResults('Hours > Group Hours')
self.groupHoursFinale('08/12/2025')
filepath = self.retrieveDownloadedFile(10, '*.html', 'Hours')
print(filepath)


  #%%      
self.kill()        

'''


'''
x = TimeClockBase(headless=True)
x.verbosity = 2
x.enableScreenshots(Path('c:/users/netadmin/downloads/test'))
x.startupBrowser()
x.tryLogin()
x.openTabularMenu()
x.searchFromTabularMenu('Group Hours')
x.clickTabularMenuSearchResults('Hours > Group Hours')
x.groupHoursFinale('05/22/2025')
filepath = x.retrieveDownloadedFile(10, '*.html', 'Hours')

'''
        
        
        
        
        
        
        
        
        