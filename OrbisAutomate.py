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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException

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

#saves list of firms to orbis profile for future access
def save_firms(browser, svname):
    
    wait = WebDriverWait(browser, 10)
    savebtn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"body > section.website > div.website__pre-content.area.area__stretchable > div > ul > li:nth-child(2) > a")))
    savebtn.click()
    
    nameenter = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"#name-field-5b34cdb74f4e40cd8c04e7095d8f3eb0")))
    nameenter.send_keys(svname)
    
    #press save
    browser.find_element_by_css_selector("body > section.website > div:nth-child(6) > div.button.popup__buttons > p > a.button.ok").click()

#Ugh so ugly. Is there a better waty to do this
def export_df(browser):
    wait = WebDriverWait(browser, 10)
    action = ActionChains(browser)
    
    results = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.results > a")))
    results.click()
    try:
        addcols = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#main-content > div > div.results > div:nth-child(2) > a")))
        addcols.click()
    except:
        addcols = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#main-content > div > div.results > div:nth-child(2) > a")))
        addcols.click()
    ###########Total assets last avail
    #Financial data 
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(6) > div"))).click()
    #Key financials
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li.with-nodes.opened.selected > ul > li:nth-child(2) > div"))).click()
    #Total assets widget
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#KEY_FINANCIALS\*KEY_FINANCIALS\.TOAS\:UNIVERSAL > div.text-container > span.px16.configuration.financialRepeat"))).click()
        
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok"))).click()
    
    ###########Operating revenue 2016
    #Operating rev widget
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#KEY_FINANCIALS\*KEY_FINANCIALS\.OPRE\:UNIVERSAL > div.text-container > span.px16.configuration.financialRepeat"))).click()
    except ElementClickInterceptedException:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#KEY_FINANCIALS\*KEY_FINANCIALS\.OPRE\:UNIVERSAL > div.text-container > span.px16.configuration.financialRepeat"))).click()
    #Change to absolute tab
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ClassicOption > div > div:nth-child(1) > div > div.list-time-selection > div.switchContainer > div > table > tbody > tr > td.right > a"))).click()
    
    
    #Check if latest avail year is already selected. If not select it.
    checkbox = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ClassicOption > div > div:nth-child(1) > div > div.list-time-selection > div.absolute-selections > div.dimension-container.yearly > div > ul > li:nth-child(6) > label")))
    if not checkbox.is_selected():
        checkbox.click()
        
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok"))).click()
    
    ###########Number of employees 2016
    search = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#Search")))
    search.send_keys("Number of Employees")
    search.send_keys(Keys.ENTER)
    
    #Num employees widget
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#KEY_FINANCIALS\*KEY_FINANCIALS\.EMPL\:UNIVERSAL > div.text-container > span:nth-child(1)"))).click()
    
    #Change to absolute tab
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ClassicOption > div > div:nth-child(1) > div > div.list-time-selection > div.switchContainer > div > table > tbody > tr > td.right > a"))).click()
        
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok"))).click()
    
    search.clear()
    
    ###########Total assets 2016
    #Key financials
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li.with-nodes.opened > ul > li:nth-child(2) > div > span"))).click()
    
    #Total assets widget
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#KEY_FINANCIALS\*KEY_FINANCIALS\.TOAS\:UNIVERSAL > div.text-container > span.px16.configuration.financialRepeat"))).click()
    
    #Change to absolute tab
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ClassicOption > div > div:nth-child(1) > div > div.list-time-selection > div.switchContainer > div > table > tbody > tr > td.right > a"))).click()
    
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok"))).click()
    
    ###########Size classification
    #LEgal section
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(5) > div"))).click()
    
    #size classification
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#LEGAL_ACCOUNT_INFORMATION\*LEGAL_ACCOUNT_INFORMATION\.COMPANY_CATEGORY\:UNIVERSAL > div.icon-container > span"))).click()
    
    ###########Addresses
    #contactinfo
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(2) > div"))).click()
    
    #Address1
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE1\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()

    #Address2
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE2\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:    
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE2\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #Address3
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE3\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:  
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE3\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #Address4
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE3\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:  
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE4\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #Postcode
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.POSTCODE\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.POSTCODE\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #city
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.CITY\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.CITY\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #std city
    try: 
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.CITY_STANDARDIZED\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.CITY_STANDARDIZED\:UNIVERSAL > div.icon-container > span"))).click()
    
    
    apply = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.list-edition-footer > form > div > input.button.ok")))
    apply.click()
    
    excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div.website__pre-content.area.area__stretchable > div.list.headerBar > div.actions > div.headerBar__exports > ul > li:nth-child(3) > a")))
    excel.click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#export-component-exportoptions > div > a > span"))).click()
    
    checkbox_srchstrat = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#export-component-exportoptions > div > div > fieldset.ExportOptions > ul > li.SearchStrategy > label")))
    checkbox_srchstrat.click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#export-component-exportoptions > div > div > fieldset.ExportOptions > ul > li:nth-child(5) > label"))).click()
    
    export = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#exportDialogForm > div.buttons.popup__buttons > a.button.submit.ok")))
    export.click()
    
    time.sleep(15)
    browser.back()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div.website__pre-content.area.area__stretchable > div > ul > li:nth-child(1) > a"))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok"))).click()
    
    
    
def orbis_auto(browser, ls):
    
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
    try:
        restartsession = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.ok")))
        restartsession.click()
    except TimeoutException:
        print("Do not need to restart orbis session")
    # time.sleep(5)

    ##################################################################################
    missed_firms = []
    ##################################################################################
    #Cycle through firm names, add to selection
    for firm in ls:  

        if browser.current_url.endswith('Report'):
            revsearch = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.search > a")))
            revsearch.click()
        else:
            browser.refresh()
        
        print(ls.index(firm))
        #Every 150 firms, manually go and download data. The webpage cannot handle ~200 or more queries in one go i think.
        if ls.index(firm)%100 == 0 and ls.index(firm) != 0:  
            export_df(browser)
            print(f"{ls.index(firm)} firms completed. File exported.")
            # save_firms(browser, svname)
            revsearch = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.search > a")))
            revsearch.click()
            
        
        orbissearch = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#search")))
        orbissearch.send_keys(f"{firm}, United States of America")
        orbissearch.send_keys(Keys.ENTER)
        # If pop-up shows up to add new search to current selection of erase current selection
        #Search for firm name, if no firm with name then move to next element in list
        try:    
            firstresult = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#quicksearch-results > ul > li:nth-child(1)")))
            firstresult.click()
        #If no search matches, continue to next name on list
        except TimeoutException:
            missed_firms.append(firm)
            continue
        time.sleep(3)

    
    return missed_firms
        
        ##################################################################################
    

if __name__ == "__main__":
    firms1 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List1_address.csv").iloc[:,0])
    firms2 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List2_address.csv").iloc[:,0])
    firms3 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List3.csv").iloc[:,0])
    firms4 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List4.csv").iloc[:,1])
    
   
    temp = firms1
    firms1_ls= []
    match = re.compile("LLC.|LLC|llc.|llc|LTD.|LTD|Ltd.|Ltd|ltd.|ltd|INC.|INC|Inc.|Inc|inc.|inc")
    for i in temp:
        try:
            pattern_split = re.split(match, i)
        except TypeError:
            continue
        l = len(pattern_split[0].split())
        i_ls = i.split()
        try:
            firms1_ls.append(pattern_split[0].replace(",", "") + i_ls[l])
        except IndexError:
            firms1_ls.append(i)
            
    firms1_ls = list(set(firms1_ls))
    
    ls = [x for x in firms1_ls if firms1_ls.index(x) > 99]
    ls1 = [x for x in ls if ls.index(x) > 102]

    browser = openbrowser()
    missed_firms2 =  orbis_auto(browser, ls)
     
    
    
    
    
    
    
    
    