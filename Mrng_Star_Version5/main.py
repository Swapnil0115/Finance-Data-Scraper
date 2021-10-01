from selenium import webdriver
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests
import lxml
import urllib.request as ur
import warnings
import re
import openpyxl
import xlsxwriter
import time
from selenium.webdriver.common.keys import Keys
import datetime
import numpy as np
from tabulate import tabulate
from varname import argname2 
from pandas import DataFrame
import inspect
from datetime import date
from selenium.webdriver.common.action_chains import ActionChains
import os
import glob
import csv
import sys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions


chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
"download.default_directory": r'C:\webdriver',
"download.prompt_for_download": False,
"download.directory_upgrade": True,
"safebrowsing.enabled": True
})
chrome_options.add_argument('--log-level=3')
chrome_options.add_argument("--output=/dev/null")
chrome_options.add_argument('--headless')
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--proxy-server='direct://'")
chrome_options.add_argument("--proxy-bypass-list=*")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--allow-running-insecure-content')
chrome_options.add_argument("--disable-in-process-stack-traces")
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument("--silent")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options)
html = driver.execute_script('return document.body.innerHTML;')
income_soup = BeautifulSoup(html,'lxml')

def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]


def Quote_Extractor_MS(Current_URL):
    driver.implicitly_wait(10)
    time.sleep(10)
    html = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html,'lxml')

    try:
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID, 'message-box-price')))
        Curr = driver.find_element_by_xpath('//*[@id="message-box-price"]').text
    except:
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'markets-components-minichart-figure.markets-components-minichart-lastClosePrice')))
        Curr = driver.find_element_by_xpath('//*[@class="markets-components-minichart-figure.markets-components-minichart-lastClosePrice"]').text

    #print(Curr)
    sal_component = []
    sal_component.append(["Last Close",Curr])

    # prev_close = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[1]/div/sal-components-mini-quote-chart/div/div[2]/div')
    # sal_component.append(prev_close.text)

    Bid_name = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[1]/div/div[1]')
    Bid = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[1]/div/div[2]')
    sal_component.append([Bid_name.text,Bid.text])

    Bid_name2 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[2]/div/div[1]')
    Bid2 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[2]/div/div[2]')
    sal_component.append([Bid_name2.text,Bid2.text])

    day_range_name = driver.find_elements_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[3]/div/div[1]')
    day_range = driver.find_elements_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[3]/div/div[2]/div')
    sal_component.append([day_range_name[0].text,day_range[0].text])

    Bid_name4 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[4]/div/div[1]')
    Bid4 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[4]/div/div[2]')
    sal_component.append([Bid_name4.text,Bid4.text])

    Bid_name5 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[5]/div/div[1]')
    Bid5 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[5]/div/div[2]')
    sal_component.append([Bid_name5.text,Bid5.text])

    Bid_name6 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[6]/div/div[1]')
    Bid6 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[6]/div/div[2]')
    sal_component.append([Bid_name6.text,Bid6.text])

    Bid_name7 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[7]/div/div[1]')
    Bid7 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[7]/div/div[2]')
    sal_component.append([Bid_name7.text,Bid7.text])

    Bid_name8 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[8]/div/div[1]')
    Bid8 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[8]/div/div[2]')
    sal_component.append([Bid_name8.text,Bid8.text])

    Bid_name9 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[9]/div/div[1]')
    Bid9 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[9]/div/div[2]')
    sal_component.append([Bid_name9.text,Bid9.text])

    Bid_name10 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[10]/div/div[1]')
    Bid10 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[10]/div/div[2]')
    sal_component.append([Bid_name10.text,Bid10.text])

    Bid_name11 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[11]/div/div[1]')
    Bid11 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[11]/div/div[2]')
    sal_component.append([Bid_name11.text,Bid11.text])

    Bid_name12 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[12]/div/div[1]')
    Bid12 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/ul/li[12]/div/div[2]')
    sal_component.append([Bid_name12.text,Bid12.text])


    

    
    sal_component_df = pd.DataFrame(sal_component,columns = ["Name","Value"])


    # divTag = income_soup.find_all("section", {"class": "sal-component-wrapper"})

    # for div in income_soup.find_all("div",{"class":"dp-value"}):
    #     sal_component.append(div.text.split())

    return sal_component_df

def KR_Extractor_MS(Current_URL):
    #driver.get(Current_URL)
    #time.sleep(10)
    KR = driver.find_element_by_xpath('//*[@class="mds-button-group"]/slot/div/mds-button[2]/label/input')
    KR.click()
    driver.implicitly_wait(10)
    time.sleep(5)
    html = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html,'lxml')

    newpage = driver.find_element_by_xpath('//*[@class="sal-full-key-ratios"]/a')

    #-----------------------------------------------------------------------START DRIVER2
    chrome_options2 = webdriver.ChromeOptions()
    chrome_options2.add_experimental_option("prefs", {
    "download.default_directory": r'C:\webdriver',
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
    })
    chrome_options2.add_argument('--log-level=3')
    chrome_options2.add_argument('--headless')
    chrome_options2.add_argument("--window-size=1920,1080")
    chrome_options2.add_argument("--disable-extensions")
    chrome_options2.add_argument("--proxy-server='direct://'")
    chrome_options2.add_argument("--proxy-bypass-list=*")
    chrome_options2.add_argument("--start-maximized")
    chrome_options2.add_argument('--disable-gpu')
    chrome_options2.add_argument('--disable-dev-shm-usage')
    chrome_options2.add_argument('--no-sandbox')
    chrome_options2.add_argument('--ignore-certificate-errors')
    chrome_options2.add_argument('--allow-running-insecure-content')
    driver2 = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options2)
    driver2.get(newpage.get_attribute('href'))
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(5)

    try:
        dfs = pd.read_html(html2)[0]
    except:
        time.sleep(5)
        dfs = pd.read_html(html2)[0]

    dfs = pd.DataFrame(dfs)

    return driver2,dfs

def Margin_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)

    dfs = pd.read_html(html2)[1]
    dfs = pd.DataFrame(dfs)
    return dfs

def Prof_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)

    dfs = pd.read_html(html2)[2]
    dfs = pd.DataFrame(dfs)

    return dfs

#Growth
def Prof2_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[2]').click()
    dfs = pd.read_html(html2)[3]
    dfs = pd.DataFrame(dfs)

    return dfs
#Cash Flow
def Prof3_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[3]').click()
    dfs = pd.read_html(html2)[4]
    dfs = pd.DataFrame(dfs)

    return dfs

#Financial Health P1
def Prof4_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[4]').click()
    dfs = pd.read_html(html2)[5]
    dfs = pd.DataFrame(dfs)

    return dfs

#Fin p2
def Prof5_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)
    dfs = pd.read_html(html2)[6]
    dfs = pd.DataFrame(dfs)

    return dfs

#Efficiency
def Prof6_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(4)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[5]').click()
    dfs = pd.read_html(html2)[7]
    dfs = pd.DataFrame(dfs)
    #driver2.close()
    return dfs

def SI_Extractor_MS(Current_URL):
    #driver.get(Current_URL)
    #time.sleep(10)
    KR = driver.find_element_by_xpath('//*[@class="mds-button-group"]/slot/div/mds-button[3]/label/input')
    KR.click()
    driver.implicitly_wait(10)
    time.sleep(5)
    html = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html,'lxml')

    sal_component = []

    # prev_close = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[1]/div/sal-components-mini-quote-chart/div/div[2]/div')
    # sal_component.append(prev_close.text)

    for i in range(1,7):
        Bid_name = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/div[2]/sal-components-short-interest/div/div[2]/ul/li['+str(i)+']/div/div[1]')
        Bid = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-quote/div/div[2]/div/div[2]/div/div[2]/div[2]/sal-components-short-interest/div/div[2]/ul/li['+str(i)+']/div/div[2]')
        sal_component.append([Bid_name.text,Bid.text])


    

    
    sal_component_df = pd.DataFrame(sal_component,columns = ["Name","Value"])


    # divTag = income_soup.find_all("section", {"class": "sal-component-wrapper"})

    # for div in income_soup.find_all("div",{"class":"dp-value"}):
    #     sal_component.append(div.text.split())

    return sal_component_df

def Summary_Extractor_MS(Current_URL):
    #driver.get(Current_URL)
    #time.sleep(10)
    driver.implicitly_wait(10)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html,'lxml')

    

    Desc = []
    Desc1 = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/section/div[2]/div/div[2]')
    Desc = Desc1.text.splitlines()
    
    Desc = Desc[:-1]
    
    Desc_df = pd.DataFrame(Desc)
    Desc_df = Desc_df.transpose()

    return Desc_df

def Description_EXT_MS(Current_URL):
    driver.implicitly_wait(10)

    Name = []
    Name.append(driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[1]/header/div/div[1]/h1').text)
    desc = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/div[1]/div/div[1]/p').text.splitlines()
    cont = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/div[1]/div/div[2]').text.splitlines()
    cont.remove('Contact')
    items = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/div[1]/div/div[3]').text.splitlines()
    
    items2 = []
    items_keys = []
    items_keys.append("Company Name")

    for i in range(len(cont)):
        items_keys.append('CONTACT'+str(i+1))
    

    for i in range(len(items)):
        if(i%2!=0):
            items2.append(items[i])
        else:
            items_keys.append(items[i])

    items_keys.append('Description')

    final = Name+cont+items2+desc
    final_df = pd.DataFrame(final).transpose()
    final_df.columns = items_keys

    return final_df

#---------------------------------------------------ANALYSIS PAGE---------------------------------
def Competitor_EXT_MS(Current_URL):
    driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/nav/ul/a[2]/span').click()
    driver.implicitly_wait(5)
    try:
        time.sleep(5)
        driver.find_elements_by_class_name('mds-button__input-outer-wrapper')[1].click()
        time.sleep(5)
        driver.find_elements_by_class_name('sal-component-expand')[0].click()
    except:
        time.sleep(15)
        driver.find_elements_by_class_name('mds-button__input-outer-wrapper')[1].click()
        time.sleep(15)
        driver.find_elements_by_class_name('sal-component-expand')[0].click()


    td_list = []

    for td in driver.find_elements_by_class_name('sal-row-title'):
        td_list.append(td.text)

    Cp_names = []
    temp = []
    for i in td_list:
        if(i=='Show Full Chart' or i=='Show Comparision Chart'):
            break
        temp.append(i)

    for i in temp:
        ind = i.find("\n")
        i = i[:ind]
        Cp_names.append(i)

    td_list = td_list[len(Cp_names)*2:]

    
    ind = td_list.index("Show Full Comparison Chart")

    row_names = td_list[ind+1:]
    td_list = td_list[:ind-1]
    td_list = list(zip(*[iter(td_list)]*len(Cp_names)))

    td_df = pd.DataFrame(td_list,columns = Cp_names)
    td_df.insert(0,"Comparision Values",row_names)


    return td_df

def News_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/news")
    driver.implicitly_wait(10)
    html = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html,'lxml')

    news = []
    links = []
    i = 1
    while(i!=200):
        try:
            news.append(driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/section/article['+str(i)+']').text.splitlines())
            news[len(news)-1].append(driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/section/article['+str(i)+']/a').get_attribute('href'))
            news[len(news)-1].remove(news[len(news)-1][0])
        except:
            break
        i+=1
        
    #print(news)

    news = pd.DataFrame(news,columns=["Article Heading","Source and Date","Article link"])    
    return news
    
def Price_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/price-fair-value")
    driver.implicitly_wait(10)
    time.sleep(10)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    dfs = pd.read_html(html)[1]
    dfs = pd.DataFrame(dfs)
    dfs = dfs.transpose()
    dfs.columns = ["Price/Fair Value","Total Return %","+/- Index"]
    dfs = dfs.drop(['Price/Fair Value'], axis = 1)
    return dfs

def Trailing_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/trailing-returns")
    driver.implicitly_wait(10)
    time.sleep(10)
    html2 = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    
    dfs2 = pd.read_html(html2)[0]
    #print(dfs2)
    return dfs2

def Trailing_EXT_MS2(driver2,Current_URL):

    driver.find_element_by_xpath('//*[@class="mds-button-group"]/slot/div/mds-button[2]').click()

    driver.implicitly_wait(10)
    time.sleep(10)
    html2 = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    
    dfs2 = pd.read_html(html2)[2]
    #print(dfs2)
    return dfs2

def Trailing_EXT_MS3(driver2,Current_URL):
    driver.find_element_by_xpath('//*[@class="mds-button-group"]/slot/div/mds-button[3]').click()
    driver.implicitly_wait(10)
    time.sleep(10)
    html2 = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    
    dfs2 = pd.read_html(html2)[4]
    #print(dfs2)
    return dfs2

#Financial Statements:-

def Financials_EXT_MS(Company_name,Current_URL):
    
    while 1:
        try:
            driver.get(Current_URL+"/"+Company_name+"/financials")
            driver.implicitly_wait(10)
            time.sleep(7)
            html = driver.execute_script('return document.body.innerHTML;')
            income_soup = BeautifulSoup(html, 'lxml')

            driver.implicitly_wait(20)
            time.sleep(5)
            Details = driver.find_element_by_xpath('//*[@class="mds-link"]')
            action = ActionChains(driver)
            action.click(on_element = Details)
            action.perform()
            break
        except:
            continue

    
    
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small')))

    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(4)  


    path = 'C:\webdriver'
    extension = 'xls'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))
    path = path+'\\'+result[0]
    DF = pd.ExcelFile(path) 
    for name in DF.sheet_names:
        df = DF.parse(name)

    DF.close()
    os.remove(path)
    time.sleep(4) 
    return df

def Financials_EXT_MS2(Company_name,Current_URL):
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    action = ActionChains(driver)
    time.sleep(5)

    Quarterly = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[2]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]')
    action.click(on_element=Quarterly)
    action.perform()

    #WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small'))).click()
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small')))
    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(4)  


    path = 'C:\webdriver'
    extension = 'xls'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))
    path = path+'\\'+result[0]
    DF = pd.ExcelFile(path) 
    for name in DF.sheet_names:
        df = DF.parse(name)

    DF.close()
    os.remove(path)
    time.sleep(4) 
    return df

#Balance Sheet
def Financials_EXT_MS3(Company_name,Current_URL):
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    action = ActionChains(driver)
    time.sleep(4)

    Bsheet = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[1]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]')
    action.click(on_element=Bsheet)
    action.perform()
    time.sleep(2)

    Annual = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[2]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[1]')
    action.click(on_element=Annual)
    action.perform()
    
    #WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small'))).click()

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small')))
    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(3)  


    path = 'C:\webdriver'
    extension = 'xls'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))
    path = path+'\\'+result[0]
    DF = pd.ExcelFile(path) 
    for name in DF.sheet_names:
        df = DF.parse(name)

    DF.close()
    os.remove(path)
    time.sleep(4) 
    return df

def Financials_EXT_MS4(Company_name,Current_URL):
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    action = ActionChains(driver)
    time.sleep(4)


    Quarter = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[2]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]')
    action.click(on_element=Quarter)
    action.perform()
    
    #WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small'))).click()

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small')))
    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(4)  


    path = 'C:\webdriver'
    extension = 'xls'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))
    path = path+'\\'+result[0]
    DF = pd.ExcelFile(path) 
    for name in DF.sheet_names:
        df = DF.parse(name)

    DF.close()
    os.remove(path)
    time.sleep(4) 
    return df

#Cash Flow:-
def Financials_EXT_MS5(Company_name,Current_URL):
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    action = ActionChains(driver)
    time.sleep(4)

    Bsheet = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[1]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[3]')
    action.click(on_element=Bsheet)
    action.perform()
    time.sleep(4)

    Annual = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[2]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[1]')
    action.click(on_element=Annual)
    action.perform()
    
    #WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small'))).click()

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small')))
    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(4)  


    path = 'C:\webdriver'
    extension = 'xls'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))
    path = path+'\\'+result[0]
    DF = pd.ExcelFile(path) 
    for name in DF.sheet_names:
        df = DF.parse(name)

    DF.close()
    os.remove(path)
    time.sleep(4) 
    return df

def Financials_EXT_MS6(Company_name,Current_URL):
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    action = ActionChains(driver)
    time.sleep(4)


    Quarter = driver.find_element_by_xpath('//*[@class="sal-component-header"]/div[2]/sal-components-segment-band/div[1]/div[1]/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]')
    action.click(on_element=Quarter)
    action.perform()

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small')))
    
    #WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small'))).click()

    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(4)  


    path = 'C:\webdriver'
    extension = 'xls'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))
    path = path+'\\'+result[0]
    DF = pd.ExcelFile(path) 
    for name in DF.sheet_names:
        df = DF.parse(name)

    DF.close()
    os.remove(path)
    time.sleep(4) 
    return df

#Valuation:-
def Valuation_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/valuation")
    driver.implicitly_wait(10)
    time.sleep(5)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    driver.find_element_by_xpath('//*[@class="sal-component-expand"]/a/span').click()
    # action = ActionChains(driver)
    # action.click(on_element = Details)
    # action.perform()

    time.sleep(5)

    dfs = pd.read_html(html)[0]
    #print(dfs)

    return dfs

#Operating performance:-
def Performance_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/performance")
    driver.implicitly_wait(10)
    time.sleep(5)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    driver.find_element_by_xpath('//*[@class="sal-component-expand"]/a/span').click()
    # action = ActionChains(driver)
    # action.click(on_element = Details)
    # action.perform()

    time.sleep(5)

    

    dfs = pd.read_html(html)[0]

    rows_to_keep = [x for x in range(dfs.shape[0]) if x not in [0,5]]
    dfs = dfs.iloc[rows_to_keep,:]
    #print(dfs)


    return dfs

#Dividends:-
def Dividends_EXT_MS(Company_name,Current_URL):
    not_avail = ['Not Available']
    try:
        driver.get(Current_URL+"/"+Company_name+"/dividends")
        driver.implicitly_wait(10)
        # driver.find_element_by_xpath('//*[@class="dividends-fixed-table"]/tbody[5]/tr[1]/td[1]/div').click()
        # driver.find_element_by_xpath('//*[@class="dividends-fixed-table"]/tbody[6]/tr[1]/td[1]/div').click()
        # driver.find_element_by_xpath('//*[@class="dividends-fixed-table"]/tbody[7]/tr[1]/td[1]/div').click()
        time.sleep(7)
        try:
            for i in range(4,8):
                driver.find_element_by_xpath('//*[@class="dividends-fixed-table"]/tbody['+str(i)+']/tr[1]/td[1]/div').click()
        except:
            try:
                for i in range(4,7):
                    driver.find_element_by_xpath('//*[@class="dividends-fixed-table"]/tbody['+str(i)+']/tr[1]/td[1]/div').click()
            except:
                try:
                    for i in range(4,6):
                        driver.find_element_by_xpath('//*[@class="dividends-fixed-table"]/tbody['+str(i)+']/tr[1]/td[1]/div').click()
                except:
                    time.sleep(5)



        time.sleep(5)
        html = driver.execute_script('return document.body.innerHTML;')

        dfs1 = pd.read_html(html)[0]
        dfs = pd.read_html(html)[2]
    except:
        dfs1 = pd.DataFrame(not_avail)
        dfs = pd.DataFrame(not_avail)


    return dfs1,dfs

def Splits_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/dividends")
    driver.implicitly_wait(10)
    time.sleep(4)
    try:
        driver.implicitly_wait(10)
        time.sleep(5)
        
        driver.find_element_by_xpath('//*[@class="mwc-tabs"]/mds-button-group/div/slot/div/mds-button[2]').click()
        time.sleep(5)
        html = driver.execute_script('return document.body.innerHTML;')

        dfs1 = pd.read_html(html)[4]
        #print(dfs1)
    except:
        not_avail = ["Not Available"]
        dfs1 = pd.DataFrame(not_avail)
        
    return dfs1

def Hist_EXT_MS(driver2,Current_URL):
    Clicker1 = driver2.find_element_by_xpath('//*[@class="r_nav"]/li[4]')
    action = ActionChains(driver2)
    action.click(on_element = Clicker1)
    action.perform()


    Clicker = driver2.find_element_by_xpath('//*[@class="r_snav"]/li[2]/a')
    action = ActionChains(driver2)
    action.click(on_element = Clicker)
    action.perform()
    time.sleep(5)

    Clicker2 = driver2.find_element_by_xpath('//*[@class="r_time_contain"]/a[7]')
    action = ActionChains(driver2)
    action.click(on_element = Clicker2)
    action.perform()
    time.sleep(5)

    driver2.implicitly_wait(10)
    Export = driver2.find_element_by_class_name('large_button_export')
    Export.click()
    time.sleep(4)  


    path = 'C:\webdriver'
    extension = 'csv'
    os.chdir(path)
    result = glob.glob('*.{}'.format(extension))

    path = path+'\\'+result[0]
    DF = pd.read_csv(path,skiprows=1)
    # for name in DF.sheet_names:
    #     df = DF.parse(name)

    # DF.close()
    os.remove(path)
    time.sleep(4) 
    #driver2.close()
    #print(DF)
    return DF

def Ownership_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/ownership")
    driver.implicitly_wait(10)

    time.sleep(6)
    html = driver.execute_script('return document.body.innerHTML;')

    Major_Funds = pd.read_html(html)[0]

    columns_to_keep = [x for x in range(Major_Funds.shape[1]) if x not in [1]]
    Major_Funds = Major_Funds.iloc[:,columns_to_keep]

    

    #Conc:-
    time.sleep(5)
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    Conc_Funds = pd.read_html(html)[0]
    columns_to_keep = [x for x in range(Conc_Funds.shape[1]) if x not in [1]]
    Conc_Funds = Conc_Funds.iloc[:,columns_to_keep]

    

    #Buying

    time.sleep(4)
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[3]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    Buy_Funds = pd.read_html(html)[0]
    columns_to_keep = [x for x in range(Buy_Funds.shape[1]) if x not in [1]]
    Buy_Funds = Buy_Funds.iloc[:,columns_to_keep]



    #Selling

    time.sleep(4)
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[4]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    Sell_Funds = pd.read_html(html)[0]
    columns_to_keep = [x for x in range(Sell_Funds.shape[1]) if x not in [1]]
    Sell_Funds = Sell_Funds.iloc[:,columns_to_keep]


    #-------------------------------------------------------Institutions----------------------------------------------------------
    driver.find_element_by_xpath('//*[@class="ownership-type-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]').click()
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[1]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')

    Major_Ins = pd.read_html(html)[0]

    columns_to_keep = [x for x in range(Major_Ins.shape[1]) if x not in [1]]
    Major_Ins = Major_Ins.iloc[:,columns_to_keep]


    #Conc:-
    time.sleep(4)
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    Conc_Ins = pd.read_html(html)[0]
    columns_to_keep = [x for x in range(Conc_Ins.shape[1]) if x not in [1]]
    Conc_Ins = Conc_Ins.iloc[:,columns_to_keep]

    #Buying

    time.sleep(4)
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[3]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    Buy_Ins = pd.read_html(html)[0]
    columns_to_keep = [x for x in range(Buy_Ins.shape[1]) if x not in [1]]
    Buy_Ins = Buy_Ins.iloc[:,columns_to_keep]


    #Selling

    time.sleep(4)
    driver.find_element_by_xpath('//*[@class="ownership-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[4]').click()
    time.sleep(4)
    html = driver.execute_script('return document.body.innerHTML;')
    Sell_Ins = pd.read_html(html)[0]
    columns_to_keep = [x for x in range(Sell_Ins.shape[1]) if x not in [1]]
    Sell_Ins = Sell_Ins.iloc[:,columns_to_keep]


    # print(Major_Funds)
    # print(Conc_Funds)
    # print(Buy_Funds)
    # print(Sell_Funds)
    # print(Major_Ins)
    # print(Conc_Ins)
    # print(Buy_Ins)
    # print(Sell_Ins)


    return Major_Funds,Conc_Funds,Buy_Funds,Sell_Funds,Major_Ins,Conc_Ins,Buy_Ins,Sell_Ins

def Execu_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/executive")
    driver.implicitly_wait(10)

    not_avail = ['Data Not available']

    time.sleep(6)
    html = driver.execute_script('return document.body.innerHTML;')

    try:
        Exec_Team = pd.read_html(html)[0]
    except:
        Exec_Team = pd.DataFrame(not_avail)
    # print(Exec_Team)

    driver.find_element_by_xpath('//*[@class="insiders-tabs"]/div/div/mwc-tabs/div/mds-button-group/div/slot/div/mds-button[2]').click()
    time.sleep(6)
    html = driver.execute_script('return document.body.innerHTML;')
    try:
        Board_Team = pd.read_html(html)[0]
    except:
        Board_Team = pd.DataFrame(not_avail)
    #print(Board_Team)

    #Transac Hist:-
    try:
        Transac= pd.read_html(html)[1]
    except:
        Transac = pd.DataFrame(not_avail)
    #print(Transac)

    return Exec_Team,Board_Team,Transac

def Execu_EXT_MS2(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/executive")
    driver.implicitly_wait(10)

    time.sleep(6)
    html = driver.execute_script('return document.body.innerHTML;')
    not_avail = ['Data Not available']
    try:
        Transac_Data = pd.read_html(html)[2]
        while(1):
            time.sleep(1)

            for i in range(1,20):
                try:
                    clicker = driver.find_element_by_xpath('//*[@class="mds-pagination"]/li['+str(i)+']/a')
                    #print(clicker.get_attribute('aria-label'))
                    if('Go to Next Page' in clicker.get_attribute('aria-label')):
                        driver.find_element_by_xpath('//*[@class="mds-pagination"]/li['+str(i)+']').click()
                        break
                    else:
                        continue
                except:
                    break


            time.sleep(1)

            check =  driver.find_element_by_xpath('//*[@class="mds-pagination"]/li[11]/a').get_attribute('class') 

            time.sleep(4)
            html_new = driver.execute_script('return document.body.innerHTML;')
            if('disabled' in check):
                break
            Transac_Data2 = pd.read_html(html_new)[2]
            time.sleep(1)
            Transac_Data = Transac_Data.append(Transac_Data2, ignore_index = True)
    except:
        Transac_Data = pd.DataFrame(not_avail)
        

    return Transac_Data

#-------------------------------------------------------------------ALL FUNCTIONS END-------------------------------------------------------------

def main(Company_name,Current_URL):
    today = date.today()
    exceldate = today.strftime("%b-%d-%Y")

    error_list = ["Cannot scrape"]

    percent_done = 0
    
    # sys.stdout.write("{:0.2f}% Completed".format(percent_done))
    #     sys.stdout.flush()
    #     time.sleep(0.05)

    # round(percent_done, 2)
    #     sys.stdout.write('\r')

    try:
        Quotes = Quote_Extractor_MS(Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Quotes = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    
    driver2,KR = KR_Extractor_MS(Current_URL)
    Margin_Sales = Margin_Extractor_MS(driver2,Current_URL)

    try:
        driver2,KR = KR_Extractor_MS(Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
        try:
            Margin_Sales = Margin_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            Margin_Sales = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

        try:
            Profitability = Prof_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            Profitability = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

        try:
            Growth = Prof2_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            Growth = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

        try:
            CashFlow = Prof3_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            CashFlow = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

        try:
            FinHealth1 = Prof4_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            FinHealth1 = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        try:
            FinHealth2 = Prof5_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            FinHealth2 = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

        try:
            Eff = Prof6_Extractor_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            Eff = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

        try:
            Historical_Data = Hist_EXT_MS(driver2,Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
        except:
            Historical_Data = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
    except:
        try:
            driver2,KR = KR_Extractor_MS(Current_URL)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)
            try:
                Margin_Sales = Margin_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                Margin_Sales = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                Profitability = Prof_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                Profitability = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                Growth = Prof2_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                Growth = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                CashFlow = Prof3_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                CashFlow = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                FinHealth1 = Prof4_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                FinHealth1 = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                FinHealth2 = Prof5_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                FinHealth2 = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                Eff = Prof6_Extractor_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                Eff = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)

            try:
                Historical_Data = Hist_EXT_MS(driver2,Current_URL)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
            except:
                Historical_Data = pd.DataFrame(error_list)
                percent_done+=3.2
                round(percent_done, 2)
                sys.stdout.write('\r')
                sys.stdout.write("{:0.2f}% Completed".format(percent_done))
                sys.stdout.flush()
                time.sleep(0.05)
        except:
            KR = pd.DataFrame(error_list)
            percent_done+=3.2
            round(percent_done, 2)
            sys.stdout.write('\r')
            sys.stdout.write("{:0.2f}% Completed".format(percent_done))
            sys.stdout.flush()
            time.sleep(0.05)

    
    try:
        Short_Interest = SI_Extractor_MS(Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Short_Interest = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Summary =  Summary_Extractor_MS(Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Summary = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Description = Description_EXT_MS(Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Description = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Competitor = Competitor_EXT_MS(Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Competitor = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        News = News_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        News = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Price_VS_Fair = Price_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Price_VS_Fair = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Trailing_Daily = Trailing_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Trailing_Daily = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Trailing_Monthly = Trailing_EXT_MS2(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Trailing_Monthly = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Trailing_Quarterly = Trailing_EXT_MS3(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Trailing_Quarterly = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    
    try:
        Income_Statement_Annual = Financials_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Income_Statement_Annual = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    
    try:
        Income_Statement_Quarterly = Financials_EXT_MS2(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Income_Statement_Quarterly = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Balance_Sheet_Annual = Financials_EXT_MS3(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Balance_Sheet_Annual = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    
    try:
        Balance_Sheet_Quarterly = Financials_EXT_MS4(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Balance_Sheet_Annual = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Cash_Flow_Annual = Financials_EXT_MS5(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Cash_Flow_Annual = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    
    try:
        Cash_Flow_Quarterly = Financials_EXT_MS6(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Cash_Flow_Quarterly = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Valuation = Valuation_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Valuation = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Performance = Performance_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Performance = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    dflist= [Quotes,Short_Interest,Summary,Description,Competitor,News,Price_VS_Fair,Trailing_Daily,Trailing_Monthly,Trailing_Quarterly,Income_Statement_Annual,Income_Statement_Quarterly,Balance_Sheet_Annual,Balance_Sheet_Quarterly,Cash_Flow_Annual,Cash_Flow_Quarterly,Valuation,Performance]

    for i in dflist:
        for col in i.columns[1:]:
            try:
                i[col] = i[col].str.replace(',', '').astype(float)
            except:
                i[col] = i[col]

      
    # We'll define an Excel writer object and the target file
    Excel_File_Name = str(exceldate) + '_' + Company_name + ".xlsx"
    Excel_File_Name = os.path.join(dir_path, Excel_File_Name)
    Excelwriter = pd.ExcelWriter(Excel_File_Name,engine="xlsxwriter",engine_kwargs={'options': {'strings_to_numbers': False}})

    KR.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=1,index=False)
    Margin_Sales.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=KR.shape[0] + 5,index=False)
    Profitability.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    Growth.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=Profitability.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    CashFlow.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=Profitability.shape[0] + 5 + Growth.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    FinHealth1.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=Profitability.shape[0] + 5 + Growth.shape[0] + 5 + CashFlow.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    FinHealth2.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=FinHealth1.shape[0] + 5 + Profitability.shape[0] + 5 + Growth.shape[0] + 5 + CashFlow.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    Eff.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=FinHealth2.shape[0] + 5 + FinHealth1.shape[0] + 5 + Profitability.shape[0] + 5 + Growth.shape[0] + 5 + CashFlow.shape[0] + 5 +KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    
    
    
    sheet_list = []
    sheet_list.append('Key_Ratios')
    #We now loop process the list of dataframes
    for df in dflist:
        sheet_list.append(retrieve_name(df)[0])
        df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],startrow=1,index=False)

        
    #--------------------------------------------------------------------DIVIDENDS---------------------------------------
    try:
        Dividends1,Dividends2 = Dividends_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Dividends1 = pd.DataFrame(error_list)
        Dividends2 = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)


    Historical_Data.to_excel(Excelwriter,sheet_name='Historical_Data',startrow=1,index=False)
    Dividends1.to_excel(Excelwriter,sheet_name='Dividends',startrow=Historical_Data.shape[0] + 5,index=False)
    Dividends2.to_excel(Excelwriter,sheet_name='Dividends',startrow=Historical_Data.shape[0] + 5 + Dividends1.shape[0] + 5,index=False)

    try:
        Splits = Splits_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Splits = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    Splits.to_excel(Excelwriter,sheet_name='Splits',startrow=1,index=False)

    try:
        Major_Funds, Concentrated_Funds, Buying_Funds, Selling_Funds, Major_Institutions, Concentrated_Institutions, Buying_Institutions, Selling_Institutions = Ownership_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Major_Funds = pd.DataFrame(error_list)
        Concentrated_Funds = pd.DataFrame(error_list)
        Buying_Funds = pd.DataFrame(error_list)
        Selling_Funds = pd.DataFrame(error_list)
        Major_Institutions = pd.DataFrame(error_list)
        Concentrated_Institutions = pd.DataFrame(error_list)
        Buying_Institutions = pd.DataFrame(error_list)
        Selling_Institutions = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Key_Executives,Board_of_Directors,Transaction_istory = Execu_EXT_MS(Company_name,Current_URL)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Key_Executives = pd.DataFrame(error_list)
        Board_of_Directors = pd.DataFrame(error_list)
        Transaction_istory = pd.DataFrame(error_list)
        percent_done+=3.2
        round(percent_done, 2)
        sys.stdout.write('\r')
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    try:
        Transaction_Hist_2 = Execu_EXT_MS2(Company_name,Current_URL)
        percent_done = 100
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)
    except:
        Transaction_Hist_2 = pd.DataFrame(error_list)
        percent_done = 100
        sys.stdout.write("{:0.2f}% Completed".format(percent_done))
        sys.stdout.flush()
        time.sleep(0.05)

    Major_Funds.to_excel(Excelwriter,sheet_name='Major_Funds',startrow=1,index=False)

    Concentrated_Funds.to_excel(Excelwriter,sheet_name='Concentrated_Funds',startrow=1,index=False)

    Buying_Funds.to_excel(Excelwriter,sheet_name='Buying_Funds',startrow=1,index=False)

    Selling_Funds.to_excel(Excelwriter,sheet_name='Selling_Funds',startrow=1,index=False)

    Major_Institutions.to_excel(Excelwriter,sheet_name='Major_Institutions',startrow=1,index=False)

    Concentrated_Institutions.to_excel(Excelwriter,sheet_name='Concentrated_Institutions',startrow=1,index=False)

    Buying_Institutions.to_excel(Excelwriter,sheet_name='Buying_Institutions',startrow=1,index=False)

    Selling_Institutions.to_excel(Excelwriter,sheet_name='Selling_Institutions',startrow=1,index=False)

    Key_Executives.to_excel(Excelwriter,sheet_name='Key_Executives',startrow=1,index=False)

    Board_of_Directors.to_excel(Excelwriter,sheet_name='Board_of_Directors',startrow=1,index=False)

    Transaction_istory.to_excel(Excelwriter,sheet_name='Transaction_istory',startrow=1,index=False)

    Transaction_Hist_2.to_excel(Excelwriter,sheet_name='Transaction_Hist_2',startrow=1,index=False)


    #34
    #--------------------------------------------------------------------DIVIDENDS---------------------------------------

    
    sheet_list.append('Dividends')
    sheet_list.append('Splits')
    sheet_list.append('Major_Funds')
    sheet_list.append('Concentrated_Funds')
    sheet_list.append('Buying_Funds')
    sheet_list.append('Selling_Funds')
    sheet_list.append('Major_Institutions')
    sheet_list.append('Concentrated_Institutions')
    sheet_list.append('Buying_Institutions')
    sheet_list.append('Selling_Institutions')
    sheet_list.append('Key_Executives')
    sheet_list.append('Board_of_Directors')
    sheet_list.append('Transaction_istory')
    sheet_list.append('Transaction_Hist_2')


    for sheet1 in sheet_list:
        # Auto-adjust columns' width
        try:
            for column in df:
                try:
                    #ExcelWriter.sheets[sheet1].write(0,column,val,header_format)
                    column_width = 20
                    col_idx = df.columns.get_loc(column)
                    Excelwriter.sheets[sheet1].set_column(col_idx, col_idx, column_width)
                except:
                    continue
        except:
            continue


    #And finally save the file
    Excelwriter.save()

    directory = "C:\webdriver"
    files_in_directory = os.listdir(directory)
    filtered_files = [file for file in files_in_directory if file.endswith(".xls")]
    for file in filtered_files:
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)
    
    filtered_files = [file for file in files_in_directory if file.endswith(".csv")]
    for file in filtered_files:
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)

    filtered_files = [file for file in files_in_directory if file.endswith(".xlsx")]
    for file in filtered_files:
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)

    

    

def test(Company_name,Current_URL):
    Key_Executives,Board_of_Directors,Transaction_istory = Execu_EXT_MS(Company_name,Current_URL)
    x2 = Execu_EXT_MS2(Company_name,Current_URL)
    
# Company_name = "ibm"
# First_Page = driver.get('https://www.morningstar.com/search?query='+Company_name)
# time.sleep(5)
# driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div/div/div[1]/div/section[1]/div[2]/a').click()
# Current_Url_ind = str(driver.current_url).find("/"+Company_name)
# Current_URL = str(driver.current_url)[:Current_Url_ind]




dire = "MorningStar"
parent_dir = "C:/"
path = os.path.join(parent_dir, dire)
if(os.path.isdir(path)):
    dir_path = path
else:
    os.mkdir(path)
    dir_path = path

def recur():
    Company_name_list = []
    while(1):
        user_input = input("ENTER TICKER NAME (TYPE 'START' AND PRESS ENTER TO STOP READING AND START EXTRACTING): ")
        if(user_input=="START"):
            break
        Company_name_list.append(user_input)


    for Company_name in Company_name_list:
        First_Page = driver.get('https://www.morningstar.com/search?query='+Company_name)
        time.sleep(5)
        driver.find_element_by_xpath('//*[@class="search-all__section"]/div[2]/a').click()
        Current_Url_ind = str(driver.current_url).find("/"+Company_name)
        Current_URL = str(driver.current_url)[:Current_Url_ind]

        # # ------------------------------------------------------test--------------------------
        # Key_Executives,Board_of_Directors,Transaction_istory = Execu_EXT_MS(Company_name,Current_URL)
        # x2 = Execu_EXT_MS2(Company_name,Current_URL)
        # # ------------------------------------------------------test--------------------------
        # Dividends_EXT_MS(Company_name,Current_URL)
        # Splits_EXT_MS(Company_name,Current_URL)
        main(Company_name,Current_URL)
        print("EXCEL FILE DOWNLOADED SUCCESSFULLY FOR --->",Company_name)
    recur()


recur()


# main(Company_name,Current_URL)

#test(Current_URL)



driver.close()

