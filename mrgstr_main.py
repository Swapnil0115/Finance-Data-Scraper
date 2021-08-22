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
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
import shutil

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
"download.default_directory": r'C:\webdriver',
"download.prompt_for_download": False,
"download.directory_upgrade": True,
"safebrowsing.enabled": True
})
driver = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=options)
html = driver.execute_script('return document.body.innerHTML;')
income_soup = BeautifulSoup(html,'lxml')

def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]


def Quote_Extractor_MS(Current_URL):
    #driver.get(Current_URL)
    #time.sleep(10)
    driver.implicitly_wait(10)
    time.sleep(10)
    html = driver.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html,'lxml')


    Curr = driver.find_element_by_xpath('//*[@id="message-box-price"]').text
    print(Curr)
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
    # options2 = webdriver.ChromeOptions()
    # options2.add_experimental_option("prefs", {
    # "download.default_directory": r'C:\webdriver',
    # "download.prompt_for_download": False,
    # "download.directory_upgrade": True,
    # "safebrowsing.enabled": True
    # })
    driver2 = webdriver.Chrome('C:\webdriver\chromedriver.exe')
    driver2.get(newpage.get_attribute('href'))
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(5)

    dfs = pd.read_html(html2)[0]
    dfs = pd.DataFrame(dfs)

    return driver2,dfs

def Margin_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)

    dfs = pd.read_html(html2)[1]
    dfs = pd.DataFrame(dfs)
    return dfs

def Prof_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)

    dfs = pd.read_html(html2)[2]
    dfs = pd.DataFrame(dfs)

    return dfs

#Growth
def Prof2_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[2]').click()
    dfs = pd.read_html(html2)[3]
    dfs = pd.DataFrame(dfs)

    return dfs
#Cash Flow
def Prof3_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[3]').click()
    dfs = pd.read_html(html2)[4]
    dfs = pd.DataFrame(dfs)

    return dfs

#Financial Health P1
def Prof4_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)

    driver2.find_element_by_xpath('//*[@id="keyStatWrap"]/div/ul/li[4]').click()
    dfs = pd.read_html(html2)[5]
    dfs = pd.DataFrame(dfs)

    return dfs

#Fin p2
def Prof5_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)
    dfs = pd.read_html(html2)[6]
    dfs = pd.DataFrame(dfs)

    return dfs

#Efficiency
def Prof6_Extractor_MS(driver2,Current_URL):
    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)

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
    time.sleep(5)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    dfs = pd.read_html(html)[1]
    dfs = pd.DataFrame(dfs)
    dfs = dfs.transpose()
    dfs.columns = ["Price/Fair Value","Total Return %","+/- Index"]
    dfs = dfs.drop(['Price/Fair Value'], axis = 1)
    return dfs

def Trailing_EXT_MS(driver2,Current_URL):
    Clicker = driver2.find_element_by_xpath('//*[@class="r_nav"]/li[4]')
    action = ActionChains(driver2)
    action.click(on_element = Clicker)
    action.perform()

    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)
    dfs2 = pd.read_html(html2)[2]
    print(dfs2)
    return dfs2

def Trailing_EXT_MS2(driver2,Current_URL):
    Clicker = driver2.find_element_by_xpath('//*[@class="in_tabs"]/li[2]')
    action = ActionChains(driver2)
    action.click(on_element = Clicker)
    action.perform()
    time.sleep(5)

    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)
    dfs2 = pd.read_html(html2)[3]
    print(dfs2)
    return dfs2

def Trailing_EXT_MS3(driver2,Current_URL):
    Clicker = driver2.find_element_by_xpath('//*[@class="in_tabs"]/li[3]')
    action = ActionChains(driver2)
    action.click(on_element = Clicker)
    action.perform()
    time.sleep(5)

    driver2.implicitly_wait(10)
    html2 = driver2.execute_script('return document.body.innerHTML;')
    income_soup = BeautifulSoup(html2,'lxml')
    time.sleep(2)
    dfs2 = pd.read_html(html2)[4]
    print(dfs2)

    driver2.close()
    return dfs2

def Financials_EXT_MS(Company_name,Current_URL):
    driver.get(Current_URL+"/"+Company_name+"/financials")
    driver.implicitly_wait(10)
    time.sleep(5)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    try:
        Details = driver.find_element_by_xpath('//*[@class="sal-summary-section"]/div[1]/a')
        action = ActionChains(driver)
        action.click(on_element = Details)
        action.perform()
        time.sleep(5)
    except:
        Details = driver.find_element_by_xpath('//*[@class="sal-summary-section"]/div[1]/a')
        Details.click()
        time.sleep(5)

    
    
    #WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sal-financials-details__export.mds-button.mds-button--small'))).click()

    Export = driver.find_element_by_class_name('sal-financials-details__export.mds-button.mds-button--small')
    Export.click()
    time.sleep(5)  


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
    return df




    



def main(Company_name,Current_URL):
    today = date.today()
    exceldate = today.strftime("%b-%d-%Y")

    error_list = ["Cannot scrape"]

    Quotes = Quote_Extractor_MS(Current_URL)
    
    driver2,KR = KR_Extractor_MS(Current_URL)

    Margin_Sales = Margin_Extractor_MS(driver2,Current_URL)

    Profitability = Prof_Extractor_MS(driver2,Current_URL)

    Growth = Prof2_Extractor_MS(driver2,Current_URL)

    CashFlow = Prof3_Extractor_MS(driver2,Current_URL)

    FinHealth1 = Prof4_Extractor_MS(driver2,Current_URL)

    FinHealth2 = Prof5_Extractor_MS(driver2,Current_URL)

    Eff = Prof6_Extractor_MS(driver2,Current_URL)

    Short_Interest = SI_Extractor_MS(Current_URL)

    Summary =  Summary_Extractor_MS(Current_URL)

    Description = Description_EXT_MS(Current_URL)

    Competitor = Competitor_EXT_MS(Current_URL)

    News = News_EXT_MS(Company_name,Current_URL)

    Price_VS_Fair = Price_EXT_MS(Company_name,Current_URL)

    Trailing_Daily = Trailing_EXT_MS(driver2,Current_URL)

    Trailing_Monthly = Trailing_EXT_MS2(driver2,Current_URL)

    Trailing_Quarterly = Trailing_EXT_MS3(driver2,Current_URL)
    
    Income_Statement_Annual = Financials_EXT_MS(Company_name,Current_URL)


    dflist= [Quotes,Short_Interest,Summary,Description,Competitor,News,Price_VS_Fair,Trailing_Daily,Trailing_Monthly,Trailing_Quarterly,Income_Statement_Annual]
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

    
    sheet_list = []
    #We now loop process the list of dataframes
    for df in dflist:
        sheet_list.append(retrieve_name(df)[0])
        df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],startrow=1,index=False)

    KR.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=1,index=False)
    Margin_Sales.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=KR.shape[0] + 5,index=False)
    Profitability.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    Growth.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=Profitability.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    CashFlow.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=Profitability.shape[0] + 5 + Growth.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    FinHealth1.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=Profitability.shape[0] + 5 + Growth.shape[0] + 5 + CashFlow.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    FinHealth2.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=FinHealth1.shape[0] + 5 + Profitability.shape[0] + 5 + Growth.shape[0] + 5 + CashFlow.shape[0] + 5 + KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    Eff.to_excel(Excelwriter,sheet_name='Key_Ratios',startrow=FinHealth2.shape[0] + 5 + FinHealth1.shape[0] + 5 + Profitability.shape[0] + 5 + Growth.shape[0] + 5 + CashFlow.shape[0] + 5 +KR.shape[0]+ 5 + Margin_Sales.shape[0] + 5,index=False)

    
    sheet_list.append('Key_Ratios')
    
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

def test(Current_URL):
    Financials_EXT_MS(Company_name,Current_URL)
    




# Company_name = input('ENTER COMPANY TICKER')
Company_name = "ibm"
First_Page = driver.get('https://www.morningstar.com/search?query='+Company_name)
time.sleep(5)
driver.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div/div/div[1]/div/section[1]/div[2]/a').click()
Current_Url_ind = str(driver.current_url).find("/"+Company_name)
Current_URL = str(driver.current_url)[:Current_Url_ind]

dir_path = os.path.dirname(os.path.realpath(__file__))

main(Company_name,Current_URL)

#test(Current_URL)



driver.close()




# #Sustainability
# Sust_Click = driver.find_elements_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/nav/ul/a[5]/span/span')[0]
# Sust_Click.click()

# driver.implicitly_wait(10)
# prev_close = driver.find_elements_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/sal-components-eqsv-sustainability/div/div[2]/div/div[2]/div/div[2]/div[1]/div[2]/sal-components-eqsv-sustainability-indicator/div/div/div[1]/div[2]/div[1]')
# print(prev_close[0].text)


# #NEWS
# news_click = driver.find_elements_by_xpath('//*[@id="__layout"]/div/div[2]/div[3]/main/nav/ul/a[3]/span/span')[0]
# news_click.click()
# driver.implicitly_wait(10)

# news_extract = driver.find_elements_by_class_name("mdc-news-module.stock__news-headline")
# for i in news_extract:
#     print("\n")
#     print(i.text)
#     print("\n")

# driver.close()


#test
# driver2 = webdriver.Chrome('C:\webdriver\chromedriver.exe')
# driver2.get('https://www.morningstar.com/stocks/xlon/ibm/price-fair-value')
# driver2.implicitly_wait(10)

# last_close = driver2.find_elements_by_class_name("legend-price")
# print(last_close[0].text)

# chart = driver2.find_elements_by_class_name("total-table")
# for i in chart:
#     print("\n")
#     print(i.text)

# driver2.close()

# driver3 = webdriver.Chrome('C:\webdriver\chromedriver.exe')
# driver3.get('https://www.morningstar.com/stocks/xlon/ibm/trailing-returns')
# driver3.implicitly_wait(10)

# table_returns = driver3.find_elements_by_class_name("sal-tab-content")
# print(table_returns[0].text)

# driver3.close()


# driver3 = webdriver.Chrome('C:\webdriver\chromedriver.exe')
# driver3.get('https://www.morningstar.com/stocks/xlon/ibm/valuation')
# driver3.implicitly_wait(10)

# table_returns = driver3.find_elements_by_class_name("mds-data-table__row__sal")
# for i in table_returns:
#     print(i.text)
#     print("\t")

# driver3.close()


# driver3 = webdriver.Chrome('C:\webdriver\chromedriver.exe')
# driver3.get('https://www.morningstar.com/stocks/xlon/ibm/ownership')
# driver3.implicitly_wait(10)

# table_returns = driver3.find_elements_by_xpath("//table/tbody/tr/td")
# table_list = []
# temp = []

# for i in table_returns:
#     temp.append(i.text)
#     if( re.findall("[a-z]\s[0-9]..\s[0-9]",temp[len(temp)-1])   ):
#         table_list.append(temp)
#         temp = []

# print(table_list)


# # starrating = driver3.find_element_by_xpath("//a[@title]")
# # print(starrating)
# html = driver3.execute_script('return document.body.innerHTML;')
# income_soup = BeautifulSoup(html, 'lxml')

# div_list = []
# for div in income_soup.find_all('title'):
#     div_list.append(div.string)
# div_list = list(filter(None, div_list))
# print(div_list)

#Remove the repeating star counts by extracting the star rating number and removing the next ((star rating number)-1) elements from array
