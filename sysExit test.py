# -*- coding: utf-8 -*-
"""
Created on Tue Feb 27 10:52:54 2018

@author: backesj
"""

import time


def yes_no(answer):
    yes = set(['yes','y'])
    no = set(['no','n',''])
     
    while True:
        choice = input(answer).lower()
        if choice in yes:
           return True
        elif choice in no:
           return False
        else:
           print("Please respond with (y/n)")

answer = yes_no(' Would you like to run the CATS report?: ')
            
if answer != True:
    raise SystemExit
else:
    time.sleep(5)
    print('program will continue')
        