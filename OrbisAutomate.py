# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 10:27:26 2021

@author: f0064wk
"""

#Dataframe stuff:
import csv
import pandas as pd
import re
import os
#Match version of chrome driver and chrome
from webdriver_manager.chrome import ChromeDriverManager
#Selenium stuff
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.proxy import Proxy, ProxyType
import selenium.webdriver.support.expected_conditions as EC
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.expected_conditions import _find_element
#Selenium excpetions
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
#Misc
import time
import os
from random import sample, seed
import numpy as np


def openbrowser():
    global browser

    url = "http://google.com/"

    capabilities = dict(DesiredCapabilities.CHROME)
    capabilities['loggingPrefs'] = { 'browser':'ALL' }

    options = webdriver.ChromeOptions()
    for arg in open('D:/git repo/scholar-scrape/config/account.txt').readlines():
        options.add_argument(arg)
    #Downloads chrome driver to match the version of chrome you use.
    browser = webdriver.Chrome(ChromeDriverManager().install(), desired_capabilities=capabilities, options=options)
    return browser

def orbis_auto(browser, ls):
    
    action = ActionChains(browser)
    
    ################################################################################
    # Login with tuck creds and remember for 30 days. Then do all this to redirect from the lubrary webpage to orbis
    browser.get("https://search.library.dartmouth.edu/discovery/search?vid=01DCL_INST:01DCL")
    
    wait = WebDriverWait(browser, 10)
    
    login = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#signInBtn")))
    login.click()
    
    loginpt2 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".md-focused")))
    loginpt2.click()
    
    time.sleep(5)
    
    browser.get("https://search.library.dartmouth.edu/discovery/fulldisplay?docid=alma991001502079705706&context=L&vid=01DCL_INST:01DCL&lang=en&search_scope=MyInst_and_CI&adaptor=Local%20Search%20Engine&tab=All&query=any,contains,orbis&offset=0")
    
    toorbis = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#getit_link1_0 > div > prm-full-view-service-container > div.section-body > div > prm-alma-viewit > prm-alma-viewit-items > md-list > div > md-list-item:nth-child(1) > div.in-element-dialog-context.layout-row.flex > div.md-list-item-text.layout-wrap.layout-row.flex > div > h5 > a")))
    toorbis.click()
    
    
    main = browser.window_handles[1]
    browser.switch_to.window(main)
    #Sometimes asks to restart session. Manually do this when waiting 15
    try:
        restartsession = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.ok")))
        restartsession.click()
    except TimeoutException:
        print("Do not need to restart")
    # time.sleep(5)

    ##################################################################################
    
    ##################################################################################
    #Cycle through firm names, add to selection
    for firm in ls:  
        orbissearch = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#search")))
        orbissearch.send_keys(f"{firm}, United States of America")
        orbissearch.send_keys(Keys.ENTER)
        # If pop-up shows up to add new search to current selection of erase current selection
        #Search for firm name, if no firm with name then move to next element in list
        try:    
            firstresult = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#quicksearch-results > ul > li:nth-child(1)")))
            firstresult.click()
        #If no search matches, continue to next name on list
        except TimeoutException:
            continue
        time.sleep(3)
        
        if ls.index(firm) == 1:
            # #Add to current selection
            # addcurrentsearch = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#rb1")))
            # addcurrentsearch.click()
            # #Do automatically for all selections
            # autodo = browser.find_element_by_css_selector("input#checkbox-50320796e2ee45a2a3f0e27783bd5113.checkbox")
            # autodo.click()
            # #click OK
            # browser.find_element_by_css_selector("body > section.website > div.popup.popup__dialog.new-search > form > div.button.popup__buttons > a.button.submit.ok").click()
            time.sleep(5)
            
        revsearch = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.search > a")))
        # action.move_to_element(revsearch).perform()
        # backtosearch = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "span.menu-tooltip")))
        revsearch.click()
        ##################################################################################
    

if __name__ == "__main__":
    firms = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List1_address.csv").iloc[:,0])
    firms = firms + list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List2_address.csv").iloc[:,0])
    firms = firms + list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List3.csv").iloc[:,0])
    firms = firms + list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List4.csv").iloc[:,1])

    
    browser = openbrowser()
    
    orbis_auto(browser, firms)
    
    
    
    
    
    
    
    