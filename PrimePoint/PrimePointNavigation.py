# -*- coding: utf-8 -*-
"""
Created on Thu Jun 11 08:51:13 2026

@author: Netadmin
"""


import time
import datetime
import os
import glob
import datetime
from PrimePoint.PrimePoint_Credentials import returnPrimePointCredentials
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException
import signal
# from psutil import Process, NoSuchProcess
import psutil
from pathlib import Path
import re 


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
    
    chrome_options.add_argument(r"--user-data-dir=C:\Users\Netadmin\AppData\Local\Google\Chrome\User Data\Profile 2")
    
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
        
        
def normalize_filename(text):
    return re.sub(r'[^a-z0-9]', '', text.lower())


#%%
class PrimePointBase():
    def __init__(self, download_folder=None, headless=False, fullscreen=False, offscreen=False):
        self.verbosity = 1
        '''
        verbosity = 0: only get messages about browser starting & download 
        verbosity = 1: get text on actions being clicked / strings entered / etc.
        verbosity = 2: get text on all element validation efforts plus verbosity = 1
        '''
        
        self.creds = returnPrimePointCredentials()
        if download_folder is None:
            download_folder = Path.home() / 'Downloads'
        # self.download_folder = download_folder
        
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
        
        self.setDownloadFolder(download_folder)
        
        if self.offscreen:
            self.moveOffscreen()
            
            
        # go fullscreen - when not headless!
        if self.fullscreen:  
            print('Browser will be FULLSCREEN')
            # full screen cus this website sucks
            self.maximizeWindow()
            
        
            
            
        # create download folder if not exists
        
        
            
            
        self.pid = self.driver.service.process.pid
        
        self.screenshotDirectory = None
        
        self.startTime = datetime.datetime.now()
        
    def setDownloadFolder(self, download_folder):
        self.download_folder = download_folder
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
       
    # def kill(self):
        
        
    #     # try:
    #     #     os.kill(int(self.pid), signal.SIGTERM)
    #     #     print(f"Killed chrome using process: {int(self.pid)}")
    #     # except ProcessLookupError as ex:
    #     #     print(ex)
    #     #     pass
    
    
        
    #     self.p = Process(self.pid)
    #     print(self.p)
    #     print(f"Killing PID {self.pid}")
        
    #     self.driver.close()
    #     print('Attempting to quit the browser...')
    #     self.driver.quit()
        
    #     # self.driver.close()
    #     print('Quitting browser completed!')
    #     try:
    #         self.p.terminate()
    #     except NoSuchProcess:
    #         print(f'Could not kill PID {self.pid} b/c it is already gone')
    
        
    
    
    def kill(self):
        """
        Gracefully shuts down Selenium/Chrome.
        Falls back to force-killing any remaining processes if necessary.
        """
    
        print("Attempting graceful browser shutdown...")
    
        try:
            self.driver.quit()
            print("driver.quit() completed")
        except Exception as ex:
            print(f"driver.quit() failed: {ex}")
    
        # Give Chrome a chance to save profile data and exit cleanly
        time.sleep(3)
    
        # Optional fallback cleanup if you store a PID
        if hasattr(self, "pid") and self.pid:
    
            try:
                if psutil.pid_exists(int(self.pid)):
    
                    print(f"PID {self.pid} still exists. Performing cleanup.")
    
                    proc = psutil.Process(int(self.pid))
    
                    # Kill child processes first
                    for child in proc.children(recursive=True):
                        try:
                            print(f"Killing child PID {child.pid}")
                            child.kill()
                        except psutil.NoSuchProcess:
                            pass
    
                    # Kill parent process
                    try:
                        proc.kill()
                        print(f"Killed PID {self.pid}")
                    except psutil.NoSuchProcess:
                        pass
    
            except Exception as ex:
                print(f"Cleanup error: {ex}")
    
        self.driver = None
    
        print("Browser shutdown complete.")    
        
        
        
    def splashPage(self):
        self.printverbosity('Navigating to splashpage!')
        # navigate to timeclock website
        self.driver.get("https://www.primepoint.net/BusinessAccess/login")
        self.printverbosity('Splashpage reached...')
        
            
    def startupBrowser(self):
        self.splashPage()
        
        # self.usernameInput = validateElement(self.driver, (By.CLASS_NAME, 'form-control'), 'usernameInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
        self.usernameInput = validateElement(self.driver, (By.CSS_SELECTOR, ".form-control:not([type='password'])"), 'usernameInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity, timeoutLimit = 30)

        self.passwordInput = validateElement(self.driver, (By.CSS_SELECTOR, ".form-control[type='password']"), 'passwordInput', checkPresence=True, checkClickable=True, verbosity=self.verbosity)
            
        
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
            
            
    def openReporting(self):
        if not self.driver.current_url == 'https://www.primepoint.net/BusinessAccess/home/dashboard':
            Exception('Cannot begin to open Reporting b/c the URL is not correct')
        
        self.ReportingButton = validateElement(self.driver, (By.LINK_TEXT, "Reporting"), 'reportingButton', True, True, verbosity=self.verbosity)
        self.ReportingButton.click()
        self.printverbosity('')
        
    

    def goToReportPage(self, reportName):
        if not self.driver.current_url == 'https://www.primepoint.net/BusinessAccess/home/reporting/interactiveReports':
            Exception('Cannot begin to open SearchBar b/c the URL is not correct')
        
        
        self.SearchBar = validateElement(self.driver, (By.XPATH, "//input[@placeholder='Search text']"), 'SearchBar',True, True, verbosity=self.verbosity)
        delete_range(self.SearchBar)
        self.SearchBar.send_keys(reportName)
        self.printverbosity('')
        
        self.Report = validateElement(self.driver, (By.XPATH, f"//div[contains(@class,'ag-cell-value') and contains(text(),'{reportName}')]"), 'Report',True, True, verbosity=self.verbosity)
        self.Report.click()    
        self.printverbosity('')
        
        self.RunReport = validateElement(self.driver, (By.XPATH, "//button[contains(@class,'btn') and contains(text(),'Run Report')]"), 'RunReport', True, True, verbosity=self.verbosity)
        self.RunReport.click()
        self.printverbosity('')


    # def downloadReportAsExcel(self):
    #     self.DownloadExcel = validateElement(self.driver, (By.XPATH, "//button[contains(@class,'btn') and contains(text(),'Download Report as Excel Spreadsheet')]"), 'DownloadExcel', True, True, verbosity=self.verbosity)
    #     self.DownloadExcel.click()
    #     self.printverbosity('')


# get both reports
# do special work after going to reportpage
# download
# move files

    def check_header(self, reportName):
        try:
            header = validateElement(self.driver, (By.XPATH, f"//span[normalize-space()='{reportName}']"), 'Header', True, True, timeoutLimit=1, verbosity=self.verbosity)
        except:
            header = None
            
        return header
    

    def getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport(self, state, date):
        
        
        reportName = 'CRYSSW Spectrum Employer Tax Expense Journal Entry Export'
        
        if self.check_header(reportName):
            self.printverbosity(f"Already on page {reportName}")
        else:
            self.goToReportPage(reportName)
            self.printverbosity(f'Navigated to report page for: {reportName}')
        
        # search format  = "{date} - {state}..."
        
        self.PayrollSelection = validateElement(self.driver, (By.TAG_NAME, "NG-SELECT-XX"), 'PayrollSelection',True, True, verbosity=self.verbosity)
        self.PayrollSelection.click()
        self.printverbosity('Clicked PayrollSelection button')
        
        self.PayrollSelectionSearchBar = validateElement(self.driver, (By.XPATH, "//input[@placeholder='Type to search...']"), 'PayrollSelectionSearchBar', True, True, verbosity=self.verbosity)
        
        
        # make sure its the right format 
        try:
            valid_date = datetime.datetime.strptime(date, '%Y-%m-%d')
        except:
            raise Exception(f"Invalid Date Format: '{date}'. Only YYYY-mm-dd allowed.")
            
        date_for_search = date
        if state == 'NY':
            valid_date = valid_date - datetime.timedelta(days=2)
            date_for_search = valid_date.strftime('%Y-%m-%d')
        # only allow fridays 
        elif valid_date.weekday() != 4:
            raise Exception(f"Provided date '{date}' is not on a Friday")
        
        # send the string we want for the date and state with an extra space
        # self.PayrollSelectionSearchBar.send_keys('{DATE AS YYYY-MM-DD} - {STATE AS 2 LETTER} ')
        # self.PayrollSelectionSearchBar.send_keys('2026-06-12 - DE ')
        search_key = f'{date_for_search} - {state} '
        self.PayrollSelectionSearchBar.send_keys(search_key)
        self.printverbosity(f"Entered '{search_key}' into PayrollSelectionSearchBar")
        # do the backspace to force a refresh
        self.PayrollSelectionSearchBar.send_keys(Keys.BACK_SPACE)
        # select it with enter
        self.PayrollSelectionSearchBar.send_keys(Keys.ENTER)
        self.printverbosity(f'Selected "{search_key}"')
        
        
        
        
        
        # files = list(Path(self.download_folder).iterdir())
        
        # if files:
        #     existing_file = max(files, key=lambda p: p.stat().st_mtime)
        #     old_mtime = existing_file.stat().st_mtime
        # else:
        #     old_mtime = 0
        
        # download
        self.startTime = datetime.datetime.now()
        self.DownloadExcel = validateElement(self.driver, (By.XPATH, "//button[contains(@class,'btn') and contains(text(),'Download Report as Excel Spreadsheet')]"), 'DownloadExcel', True, True, verbosity=self.verbosity)
        self.DownloadExcel.click()
        self.printverbosity('Clicked Download Excel button')
        
        
        
        
        # filepath = self.retrieveDownloadedFile(waitTime=10, fileType='*.xlsx', searchText=reportName)
        
        filepath = self.retrieveDownloadedFile(waitTime=10, searchText=reportName)
        
        self.printverbosity(f'File Downloaded as: {filepath}')
        
        
        date_file_format = valid_date.strftime('%m.%d.%Y')
        filepath_pathlib = Path(filepath)
        renamed_file = f'{state} {date_file_format} BS {self.startTime.strftime("%Y%m%d_%H%M%S")}{filepath_pathlib.suffix}'
        
        
        new_filepath = filepath_pathlib.with_name(renamed_file)
        if os.path.exists(new_filepath):
            self.printverbosity(f'Overwriting file because it exists: {new_filepath}')
            
              
        
        filepath_pathlib.replace(new_filepath)
        # os.remove(filepath_pathlib)
        
        self.printverbosity(f'File renamed to: {new_filepath}')
        
        if os.path.exists(new_filepath):
            return new_filepath
        else:
            raise Exception(f'Did not successfully create new file: {new_filepath}')
        
                
                      
            
    def getLaborDistributionPayrollSummary(self, state, date):
        
        
        
        reportName = 'Labor Distribution Payroll Summary'
        
        if self.check_header(reportName):
            self.printverbosity(f"Already on page {reportName}")
        else:
            self.goToReportPage(reportName)
            self.printverbosity(f'Navigated to report page for: {reportName}')

        self.PayrollSelection = validateElement(self.driver, (By.TAG_NAME, "NG-SELECT-XX"), 'PayrollSelection',True, True, verbosity=self.verbosity)
        self.PayrollSelection.click()
        self.printverbosity('Clicked PayrollSelection')
        self.PayrollSelectionSearchBar = validateElement(self.driver, (By.XPATH, "//input[@placeholder='Type to search...']"), 'PayrollSelectionSearchBar', True, True, verbosity=self.verbosity)
        

        # make sure its the right format 
        try:
            valid_date = datetime.datetime.strptime(date, '%Y-%m-%d')
        except:
            raise Exception(f"Invalid Date Format: '{date}'. Only YYYY-mm-dd allowed.")
            
        date_for_search = date
        if state == 'NY':
            valid_date = valid_date - datetime.timedelta(days=2)
            date_for_search = valid_date.strftime('%Y-%m-%d')
        # only allow fridays 
        elif valid_date.weekday() != 4:
            raise Exception(f"Provided date '{date}' is not on a Friday")
        
        # send the string we want for the date and state with an extra space
        # self.PayrollSelectionSearchBar.send_keys('{DATE AS YYYY-MM-DD} - {STATE AS 2 LETTER} ')
        # self.PayrollSelectionSearchBar.send_keys('2026-06-12 - DE ')
        search_key = f'{date_for_search} - {state} '
        
        self.PayrollSelectionSearchBar.send_keys(search_key)
        self.printverbosity(f"Entered {search_key} into PayrollSelectionSearchBar")
        # do the backspace to force a refresh
        self.PayrollSelectionSearchBar.send_keys(Keys.BACK_SPACE)
        # select it with enter
        self.PayrollSelectionSearchBar.send_keys(Keys.ENTER)
        self.printverbosity(f"Selected '{search_key}'")
        
        
        # self.AllocationLevelDropdown = validateElement(self.driver, (By.TAG_NAME, "SELECT"), "AllocationLevelDropdown", True, True, verbosity=self.verbosity)
        self.AllocationLevelDropdown = validateElement(self.driver, (By.XPATH, "//label[@for='ALLOC_LEVEL_1']/following-sibling::div//select"), 'FileFormatDropdown', True, True, verbosity=self.verbosity)
        Select(self.AllocationLevelDropdown).select_by_visible_text("-none-")
        self.printverbosity('Selected "-none-" from AllocationLevelDropdown')

        # ENABLE ONLY DISPLAY SUMMARY PAGE
        self.OnlyShowSummaryPageCheckbox = validateElement(self.driver, (By.XPATH, "//label[@for='ONLY_SHOW_SUMMARY_PAGE']/following-sibling::div//input[@type='checkbox']"), 'OnlyShowSummaryPageCheckbox', True, True, verbosity=self.verbosity)
        if not self.OnlyShowSummaryPageCheckbox.is_selected():
            self.OnlyShowSummaryPageCheckbox.click()
            self.printverbosity(f'OnlyShowSummaryPageCheckbox was unchecked and is now: {self.OnlyShowSummaryPageCheckbox.is_selected()}')
            
            
        
        
        # DISABLE SHOW SUMMARY PAGE
        self.ShowSummaryPageCheckbox = validateElement(self.driver, (By.XPATH, "//label[@for='SHOW_SUMMARY_PAGE']/following-sibling::div//input[@type='checkbox']"), 'ShowSummaryPageCheckbox', True, True, verbosity=self.verbosity)
        if self.ShowSummaryPageCheckbox.is_selected():
            self.ShowSummaryPageCheckbox.click()
            self.printverbosity(f'ShowSummaryPageCheckbox was unchecked and is now: {self.ShowSummaryPageCheckbox.is_selected()}')
            
            
            
        self.FileFormatDropdown = validateElement(self.driver, (By.XPATH, "//label[@for='fileFormat']/following-sibling::div//select"), 'FileFormatDropdown', True, True, verbosity=self.verbosity)
        Select(self.FileFormatDropdown).select_by_visible_text("Excel")
        self.printverbosity('Selected "Excel" from FileFormatDropdown')

        
        self.startTime = datetime.datetime.now()
        self.DownloadButton = validateElement(self.driver, (By.XPATH, "//button[contains(normalize-space(), 'Download Report')]"), "DownloadButton", True, True, verbosity=self.verbosity)
        self.DownloadButton.click()       
        self.printverbosity('DownloadButton Clicked')
        
        # filepath = self.retrieveDownloadedFile(waitTime=10, fileType='*.xls', searchText=reportName.replace(' ' , '_'))
        filepath = self.retrieveDownloadedFile(waitTime=10, searchText=reportName)
        
        self.printverbosity(f'File Downloaded as: {filepath}')
        
        date_file_format = valid_date.strftime('%m.%d.%Y')
        filepath_pathlib = Path(filepath)
        renamed_file = f'{state} {date_file_format} Tax Report {self.startTime.strftime("%Y%m%d_%H%M%S")}{filepath_pathlib.suffix}'
        new_filepath = filepath_pathlib.with_name(renamed_file)
        filepath_pathlib.replace(new_filepath)
        
        self.printverbosity(f'File renamed to: {new_filepath}')
        
        if os.path.exists(new_filepath):
            return new_filepath
        else:
            raise Exception(f'Did not successfully create new file: {new_filepath}')
        
        
    def goToReportingFromSpecificReportPage(self):
        try:
            # this is the back button for getLaborDistributionPayrollSummary
            self.BackToAllReportsButton = validateElement(self.driver, (By.XPATH, "//button[contains(normalize-space(), 'Back to All Reports')]"), "BackToAllReportsButton", True, True, timeoutLimit=1, verbosity=self.verbosity)
            self.BackToAllReportsButton.click()   
            
        except UnboundLocalError:
            try:
                # this is the back button for getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport
                self.CloseButton = validateElement(self.driver, (By.XPATH, "//button[contains(normalize-space(), 'Close')]"), "CloseButton", True, True, timeoutLimit=1, verbosity=self.verbosity)
                self.CloseButton.click()   
            except UnboundLocalError:
                raise Exception('Could not go back to Reporting page. Try starting over!')
  
        
        
    def retrieveDownloadedFile(self, waitTime, searchText):
        
        normalized_search = normalize_filename(searchText)
    
        matching_files = [f for f in Path(self.download_folder).iterdir() if f.is_file() and normalized_search in normalize_filename(f.name)]
    
        old_mtime = 0
    
        if matching_files:
            old_mtime = max(f.stat().st_mtime for f in matching_files)
    
        self.printverbosity(f"Waiting for download containing '{searchText}'")
    
        endTime = time.time() + waitTime
    
        while time.time() < endTime:
    
            matching_files = [
                f for f in Path(self.download_folder).iterdir()
                if (
                    f.is_file()
                    and normalized_search in normalize_filename(f.name)
                    and not f.name.endswith(".crdownload")
                )
            ]
    
            if matching_files:
    
                newest = max(matching_files, key=lambda p: p.stat().st_mtime)
    
                if newest.stat().st_mtime > old_mtime:
    
                    # Wait until file size stabilizes
                    size1 = newest.stat().st_size
                    time.sleep(0.25)
                    size2 = newest.stat().st_size
    
                    if size1 == size2:
                        self.printverbosity(f"File found: {newest}")
                        return str(newest)
    
            time.sleep(0.25)
    
        raise WhileTimerTimeoutExcpetion(f"No downloaded file containing '{searchText}' found.")
    
                
        
class PrimePointEZ(PrimePointBase):
    def __init__(self, friday_date_string, offscreen=True, headless=False):
        
        
        self.friday_date_string = friday_date_string
        self.offscreen = offscreen
        self.headless = headless
        self.debug = False
        
        
        self.download_folder = Path.home() / 'Downloads' / 'PRIMEPOINT'
        
        self.states_list = ['MD','DE','TN','PA','NY']
        self.filepaths_to_move = []
 
    def get_filepaths_wrapper(self):
        try:
            self.get_filepaths()
        finally:
            self.kill()
 
    def get_filepaths(self):
        if self.offscreen:
            self.tcb = PrimePointBase(download_folder=self.download_folder, offscreen=self.offscreen, fullscreen=True)
            
        
        elif self.headless:
            self.tcb = PrimePointBase(download_folder=self.download_folder, headless=True)
        
        self.tcb.verbosity=1
        
        if self.debug:
            self.tcb.verbosity = 2
            self.tcb.screenshotDirectory = 'C:\\Users\\Netadmin\\Downloads\\New folder'
        
        self.tcb.startupBrowser()
        self.tcb.tryLogin()
        
        self.tcb.openReporting()
        # get all of the CRYSSW xxxx reports 
        self.tcb.setDownloadFolder(self.download_folder / 'getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport')
        for state in self.states_list:
            filepath = self.tcb.getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport(state=state, date=self.friday_date_string)
            self.filepaths_to_move.append(filepath)

        # go back to the list of available reports
        self.tcb.goToReportingFromSpecificReportPage()

        
        # get all of the Labor Distribution reports
        self.tcb.setDownloadFolder(self.download_folder / 'getLaborDistributionPayrollSummary')
        for state in self.states_list:
            filepath = self.tcb.getLaborDistributionPayrollSummary(state=state, date=self.friday_date_string)
            self.filepaths_to_move.append(filepath)
        
        
    
    def kill(self):
        self.tcb.kill()
        
        
'''

y = PrimePointEZ('2026-06-12')
y.get_filepaths_wrapper()
y.kill()


'''        
        
        


'''
x = PrimePointBase(headless=False)
x.verbosity = 0
# x.enableScreenshots(Path('c:/users/netadmin/downloads/test'))
x.startupBrowser()
x.tryLogin()
x.openReporting()
# x.setDownloadFolder(Path.home() / 'Downloads' / 'getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport')
# filepath = x.getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport(state='DE', date='2026-06-12')
# filepath = x.getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport(state='TN', date='2026-06-12')
# filepath = x.getCRYSSWSpectrumEmployerTaxExpenseJournalEntryExport(state='NY', date='2026-06-12')



# x.goToReportingFromSpecificReportPage()
x.setDownloadFolder(Path.home() / 'Downloads' / 'getLaborDistributionPayrollSummary')

filepath1 = x.getLaborDistributionPayrollSummary('DE', '2026-06-12')
filepath1 = x.getLaborDistributionPayrollSummary('TN', '2026-06-12')
filepath1 = x.getLaborDistributionPayrollSummary('PA', '2026-06-12')

x.kill()
'''
        
        
        
        
        
        
        
        
        