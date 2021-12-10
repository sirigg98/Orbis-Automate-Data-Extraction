# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 18:23:06 2021

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
import selenium.webdriver.support.expected_conditions as EC
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.expected_conditions import _find_element
#Selenium excpetions
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException, ElementClickInterceptedException
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

def search_cap_IQ(browser,wait,  name):
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#_CPH__companyOptionsSection__companyLookup_companyLookupDS_companyNameSearchRadioBtn"))).click()
    except ElementClickInterceptedException:
        return
    searchbar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#_CPH__companyOptionsSection__companyLookup_companyLookupDS_companyNameSearchTextbox")))
    searchbar.send_keys(str(name))
    searchbar.send_keys(Keys.ENTER)
    searchbar.clear()
    
    try:
        firstresult = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#_CPH__companyOptionsSection__companyLookup_companyLookupDS_companySelectorLBCC_optionsList > option:nth-child(1)")))
        firstresult.click()
    except (TimeoutException, StaleElementReferenceException) as e:
        return
    
    # #Checking if company profile. sometimes it catches a professional profile
    # try:
    #     addtolist = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#CompanyHeaderInfo__AddToListIcon__AddToListIcon_AddToListImageMap > area:nth-child(2)")))
    #     addtolist.click()
    # except TimeoutException: 
    #     print(str(name) + "'s page has no add to list button")
    #     return
    prelimaddbtn = browser.find_element_by_css_selector("#_CPH__companyOptionsSection__companyLookup_companyLookupDS_companySelectorLBCC_addBtn")
    ActionChains(browser).move_to_element(prelimaddbtn).click().perform()

    
    # compset = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#CompanyHeaderInfo__AddToListFadePanelMenu_Body > table > tbody > tr:nth-child(4) > td > a > span")))
    # compset.click()
    
    # iframe = browser.find_element_by_id("addToListIFrame")
    # browser.switch_to.frame(iframe)
    
    # savebtn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#_AddToListControl__lnms_CompanyCompSetMsi_CompanyCompSetLms_CompanyCompSetListSelector_ctl03_ctl01__saveBtn")))
    # ActionChains(browser).move_to_element(savebtn).click().perform()
    
    # browser.get("https://www.capitaliq.com/CIQDotNet/my/dashboard.aspx")
   
    
# firms1 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List1_address.csv").iloc[:,0])
firms2 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List2_address.csv").iloc[:,0])
# firms3 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List3.csv").iloc[:,0])
# firms4 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List4.csv").iloc[:,1])

   
temp = firms2
firms2_ls= []
match = re.compile("LLC.|LLC|llc.|llc|LTD.|LTD|Ltd.|Ltd|ltd.|ltd|INC.|INC|Inc.|Inc|inc.|inc")
for i in temp:
    pattern_split = re.split(match, i)
    l = len(pattern_split[0].split())
    i_ls = i.split()
    try:
        firms2_ls.append(pattern_split[0].replace(",", "") + i_ls[l])
    except IndexError:
        firms2_ls.append(i)    




browser = openbrowser()
wait = WebDriverWait(browser, 5)
browser.get("https://www.capitaliq.com/CIQDotNet/Comps/Comparables.aspx?objectId=1196252625&statekey=aeaa53600f344f9d8db4c7871fe15669")

browser.find_element_by_css_selector("#myLoginButton").click()

for i in firms2_ls:
    print(firms2_ls.index(i))
    search_cap_IQ(browser, wait, i)
