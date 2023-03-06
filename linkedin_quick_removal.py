# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 11:36:16 2020

@author: CWilson
"""

''' linkedin deleter '''


import pyautogui as pag
import keyboard
import time



# first you need to define where the final button is
# press the '-' when hovering the mouse over the remove button
print('press the minus key "-" to set location of remove button')
while True:
    if keyboard.is_pressed('-'):
        remove_button = pag.position()
        pag.click()
        break
    



print('You may now begin removing connections')
print('press the minus key "-" while hovering over the "..." to remove connection')
print('press the equal key "=" to quit')

count = 0
while True:
    if keyboard.is_pressed('-'):
        start = pag.position()
        pag.click()
        pag.moveRel(0,45)
        pag.click()
        pag.moveTo(remove_button)
        pag.click()
        count += 1
        pag.moveTo(start)
    
    if keyboard.is_pressed('='):
        break
    
    
    
print('removed about {} connections'.format(count))