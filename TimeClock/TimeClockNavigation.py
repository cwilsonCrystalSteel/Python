# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:36:05 2024

@author: CWilson
"""

import time
import datetime
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from TimeClock_Credentials import returnTimeClockCredentials
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


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


#%%
class TimeClockBase():
    def __init__(self, download_folder='c:\\users\\cwilson\\downloads', headless=True, fullscreen=False):
        
        self.creds = returnTimeClockCredentials()
        self.download_folder = download_folder
        self.headless = headless
        self.fullscreen = fullscreen
        
        
            
            
        if self.fullscreen and self.headless:
            Exception('Both options fullscreen & headless are enabled, only one can be enabled at a time')
            
        self.service = Service()
            
            
        if self.headless:
            print('Browser will be HEADLESS')
            options = setting_chrome_options(True)
        else:
            options = setting_chrome_options(False)
            
            
        self.driver = webdriver.Chrome(service=self.service, options = options)
            
            
        if self.fullscreen:  
            print('Browser will be FULLSCREEN')
            # full screen cus this website sucks
            self.driver.maximize_window()
            
        
        
        
        
        if os.path.exists(download_folder):
            enable_download(self.driver, self.download_folder)
        else:
            os.makedirs(download_folder)
            enable_download(self.driver, self.download_folder)
            
        print('Driver created...')
        
    def kill(self):
        self.driver.quit()
        
    def splashPage(self):
        # navigate to timeclock website
        self.driver.get("https://136509.tcplusondemand.com/app/manager/#/ManagerLogOn/136509")
        
            
    def startupBrowser(self):

                

        
        # printwait('Loading website', 6)
        try:
            self.userid = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'LogOnUserId')))
            print('The login page was loaded in a timely manner - username box')
        except TimeoutException:
            print('Loading the login page took too long!')
            
        try:
            self.password = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'LogOnUserPassword')))
            print('The login page was loaded in a timely manner - password box')
        except TimeoutException:
            print('Loading the login page took too long!')      
        
        
    def login(self):
        # userid = self.driver.find_element(By.ID, 'LogOnUserId')
        # get the username & password dict
        self.userid.send_keys(self.creds['username'])
        printwait('Entered Username', 0)
        # Find the password field
        # password = self.driver.find_element(By.ID, 'LogOnUserPassword')
        # Submit password
        self.password.send_keys(self.creds['password'])
        printwait('Entered Password', 0)
        # Press 'enter' to login
        self.password.send_keys(Keys.RETURN)
        
    def openTabularMenu_InitiateSearch(self, searchText=''):
        
        try:
            self.tabularMenu = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'HeaderMenuIcon')))
            print('Found Tabular Menu')
        except TimeoutException:
            print('Timeout Exception for the Tabular Menu') 
        
        # self.tabularmenu = self.driver.find_element(By.CLASS_NAME, 'HeaderMenuIcon')
        # printwait('Found Tabular menu', 3) 
        self.tabularMenu.click()
        
        try:
            self.tabularMenuSearchBox = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'SearchInput')))
            print('Found Search Box')
        except TimeoutException:
            print('Timeout Exception for the Tabular Menu Search Box') 
        
        
        # printwait('Found search box', 3) 
        # self.searchbox = self.driver.find_element(By.CLASS_NAME, 'SearchInput')
        self.tabularMenuSearchBox.send_keys(searchText)
        self.tabularMenuSearchBox.send_keys(Keys.RETURN)
        printwait('Searched "Group Hours"', 0) 
        
        
        
        
        
        
        
x = TimeClockBase(headless=True)     
x.splashPage()  
x.startupBrowser()
x.login()
x.openTabularMenu_InitiateSearch('Group Hours')

  #%%      
        

        
        
        
        
        
        
        
        
        
        