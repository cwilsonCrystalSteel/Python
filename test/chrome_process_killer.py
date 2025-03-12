# -*- coding: utf-8 -*-
"""
Created on Fri Oct 11 16:31:13 2024

@author: CWilson
"""

import psutil
import time

current_time = time.time()

for proc in psutil.process_iter(['pid','name','create_time','cmdline']):
    try:
        if proc.info['name'] and 'chrome' in proc.info['name'].lower():
        
            process_age_seconds = current_time - proc.info['create_time']
            if process_age_seconds > 2*60:
                proc.kill()
                print(f"Killed process: {proc.info['name']} (PID: {proc.info['pid']})")
                if proc.info['cmdline'] is not None:
                    print(proc.info['cmdline'])
                    
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
        print(e)
        print(f"Failed to kill: {proc.info['name']} (PID: {proc.info['pid']})")
        pass
    