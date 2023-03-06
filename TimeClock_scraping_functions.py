# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 09:00:47 2021

@author: CWilson
"""
import time
import datetime
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

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


def setting_chrome_options():
    chrome_options = Options()
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