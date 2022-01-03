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
from selenium.common.exceptions import ElementNotInteractableException, NoSuchElementException, TimeoutException, ElementClickInterceptedException, StaleElementReferenceException
#Misc
from difflib import SequenceMatcher 
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
    wait = WebDriverWait(browser, 60)
    
    results = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.results > a")))
    time.sleep(5)
    results.click()
    
    time.sleep(40)
    try:
        addcols = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#main-content > div > div.results > div:nth-child(2) > a")))
        addcols.click()
    except ElementClickInterceptedException:
        addcols = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#main-content > div > div.results > div:nth-child(2) > a")))
        addcols.click()
    ###########Total assets last avail
    #Financial data 
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(6) > div"))).click()
    #Key financials
    wait.untilcd(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li.with-nodes.opened.selected > ul > li:nth-child(2) > div"))).click()
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
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(5) > div"))).click()
    except ElementClickInterceptedException:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(5) > div"))).click()
    #size classification
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#LEGAL_ACCOUNT_INFORMATION\*LEGAL_ACCOUNT_INFORMATION\.COMPANY_CATEGORY\:UNIVERSAL > div.icon-container > span"))).click()
    
    ########### Industry Classification
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(3) > div"))).click()
    time.sleep(2)
    
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li.with-nodes.selected.opened > ul > li:nth-child(1) > div"))).click()
    time.sleep(2)
    naics = browser.find_element_by_css_selector("#INDUSTRY_ACTIVITIES\*INDUSTRY_ACTIVITIES\.NAICS2017_CORE_CODE\:UNIVERSAL > div.icon-container > span")
    actions = ActionChains(browser)
    actions.move_to_element(naics).perform()
    naics.click()
    
    # try:
    #     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#INDUSTRY_ACTIVITIES\*INDUSTRY_ACTIVITIES\.NAICS2017_CORE_CODE\:UNIVERSAL > div.icon-container > span"))).click()
    # except TimeoutException:
    #     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#INDUSTRY_ACTIVITIES\*INDUSTRY_ACTIVITIES\.NAICS2017_CORE_CODE\:UNIVERSAL > div.icon-container > span"))).click()


    
    ###########Addresses
    #contactinfo
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div.format-editor-widget > div:nth-child(1) > div > div:nth-child(2) > div > ul > li:nth-child(2) > div"))).click()
    


    
    #Postcode
    try:
        zipcode = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.POSTCODE\:UNIVERSAL > div.icon-container > span")))
    except ElementClickInterceptedException:
        zipcode = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.POSTCODE\:UNIVERSAL > div.icon-container > span")))
    
    actions = ActionChains(browser)
    actions.move_to_element(zipcode).perform()
    zipcode.click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #Address1
    addy1 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE1\:UNIVERSAL > div.icon-container > span")))
    actions.move_to_element(addy1).perform()
    addy1.click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()

    #Address2
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE2\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:    
        time.sleep(2)
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
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE4\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:  
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.ADDRESS_LINE4\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
    #city
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.CITY\:UNIVERSAL > div.icon-container > span"))).click()
    except ElementClickInterceptedException:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#CONTACT_INFORMATION\*CONTACT_INFORMATION\.CITY\:UNIVERSAL > div.icon-container > span"))).click()
    save = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.website > div:nth-child(6) > div.buttons.popup__buttons > a.button.ok")))
    save.click()
    
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
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > section.side-main > div > ul > li.search > a"))).click()
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
    matched_firms = []
    ##################################################################################
    #Cycle through firm names, add to selection
    
    for firm in ls:   
            
        browser.refresh()
        orbissearch = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#search")))
        orbissearch.send_keys(f"{firm[1]}, United States of America")
        orbissearch.send_keys(Keys.ENTER)
        
    
        print(ls.index(firm))
        # If pop-up shows up to add new search to current selection of erase current selection
        #Search for firm name, if no firm with name then move to next element in list
        try:    
            firstresult = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#quicksearch-results > ul > li:nth-child(1)")))
            orbisname = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#quicksearch-results > ul > li:nth-child(1) > a > div.column.column-flex > div.column.column-name > p.name"))).text
            matched_firms.append([orbisname, firm[0], firm[1]])
            firstresult.click()
            
        #If no search matches, continue to next name on list
        except TimeoutException:
            matched_firms.append(["", firm[0], firm[1]])
            continue
        
        if ls.index(firm)%100 == 0 or ls.index(firm) == len(ls)-1:
            if ls.index(firm) != 0:
                print('Exporting...')
                export_df(browser)
                print(f"{ls.index(firm)} firms completed. File exported.")
            # save_firms(browser, svname)
                revsearch = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.search > a")))
                revsearch.click()
        
        time.sleep(3)    
            
        if "report" in browser.current_url.lower():
            revsearch = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"body > section.side-main > div > ul > li.search > a")))
            revsearch.click()
            time.sleep(3)
    return matched_firms

def replace_all(text, dic):
    for i, j in dic.items():
        text = text.replace(i, j)
    return text
        
        ##################################################################################
    

if __name__ == "__main__":
    firms1 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List1_address.csv").iloc[:,0])
    firms2 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List2_address.csv").iloc[:,0])
    firms3 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List3.csv").iloc[:,0])
    firms3_2 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\CheckList3_2.csv").iloc[:,1])
    firms4 = list(pd.read_csv(r"C:\Users\F0064WK\OneDrive - Tuck School of Business at Dartmouth College\Documents\US Tariff Exemptions\USTR data\Check_List4.csv").iloc[:,1])
    
   
    temp = [x for x in firms1 if x != "E  E CO LTD DBA JLA HOME"]
    temp = list(set(temp))
    firms3_ls= []
    match = re.compile("llc.|llc|ltd.|ltd|inc.|inc|corp.|corporation|corp|company|companies|co.")
    match_dict = {" llc.": "",
                  " llc": "",
                  " ltd.": "",
                  " ltd": "",
                  " limited": "",
                  " incorporated":"",
                  " inc.":"",
                  " inc":"",
                  " corp.":"",
                  " corporation":"",
                  " corp":"",
                  " company":"",
                  " companies":"",
                  " co.": "",
                  " co": "",
                  ",": "", 
                  "united states of america": "",
                  "usa": "", 
                  "(":"", 
                  ")": "", 
                  " u.s.a":"", 
                  " +":"", 
                  " america" : ""}
    
    firmnames = []
    for i in temp:
        if pd.isna(i):
            i = ''
        if "Continental" in i:
            firmnames.append([i, i.replace(",", "").replace(".", "")])
            continue
        try:
            pattern_split = re.split(match, i.lower())
        except TypeError:
            continue
        l = len(pattern_split[0].replace(",", "").replace(".", "").split())
        i_ls = i.replace(",", "").replace(".", "").split()
        if "dba" in i:
            i_split = i.split("dba")
            i = i_split[0]
            continue
        if "d/b/a" in i:
            i_split = i.split("d/b/a")
            i = i_split[0]
            continue
        try:
            firmnames.append([i, pattern_split[0] + i_ls[int(l+1)]])
        except IndexError:
            firmnames.append([i, i])
            
    firmnames = [x for x in firmnames if firmnames.index(x)> 302]
    # firms_orbis = list(pd.read_csv(r"C:\Users\F0064WK\Downloads\CheckList2_firmsize_address.csv").iloc[:,0])
    # matched_names = []
    
    # for i in USTRfirms:
    #     print(USTRfirms.index(i))
    #     orbisname = replace_all(i[0].lower(), match_dict)
    #     for j in firms_orbis:
    #         USTRname = replace_all(j.lower(), match_dict)
    #         USTRname_ls = USTRname.split()
    #         ratio = SequenceMatcher(None, orbisname, USTRname).ratio()
    #         matched_names.append([j, i[0], i[1] + ", United States of America", ratio])
    #         if ratio == 1:
    #             firms_orbis.remove(j)
        
        
    # match_names_df_temp = pd.DataFrame(matched_names, columns = ["Orbis Name", "USTR Name", "Search Term", "Fuzzy_Ratio"])
    # match_names_df = match_names_df_temp.sort_values("Fuzzy_Ratio", ascending = False).drop_duplicates("Orbis Name").set_index("Orbis Name")

    # match_names_df.to_csv(r"C:\Users\F0064WK\Downloads\names2.csv")
    # ls = [x for x in firms3_2_ls if firms3_2_ls.index(x) > 159
    
    browser = openbrowser()
    # matched_firms =  orbis_auto(browser, firmnames)
    # for i in matched_firms:
    #     i.append(SequenceMatcher(None, i[1].lower(), i[0].lower()).ratio())
    # df = pd.DataFrame(matched_firms, columns=['orbis_name', 'ustr_name', 'search_term', 'match_ratio'])
    # df.to_csv(r'C:\Users\F0064WK\Downloads\crosswalk_1.csv')
    # #export_df(browser)
