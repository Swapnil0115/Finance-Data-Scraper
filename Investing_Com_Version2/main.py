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
# chrome_options.add_argument('--headless')
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
chrome_options.add_experimental_option("excludeSwitches", ["disable-popup-blocking"])

driver = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options)


def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]




def Summary_iv():
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')
    
    time.sleep(5)

    Summary = []
    
    for i,j in zip(soup.find_all("dt"),soup.find_all("dd")):
        Summary.append([i.text,j.text])

    Summary = pd.DataFrame(Summary,columns = ["Key","Value"])
    
    return Summary


def Profile_iv():
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'navbar_navbar__2yeca')))
    driver.find_element_by_xpath('//*[@class="navbar_navbar__2yeca"]/ul/li[2]').click()
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')


    Name = ['Company Name']

    Prof_Body = ['Description']

    Name.append(driver.find_element_by_class_name('instrumentHeader').text.splitlines()[0])

    Prof = driver.find_element_by_class_name('companyProfileHeader').text.splitlines()  

    Prof_Body.append(driver.find_element_by_class_name('companyProfileBody').text.splitlines()[0])

    Contact = driver.find_element_by_class_name('companyProfileContactInfo').text.splitlines()

    Prof = list(zip(*[iter(Prof)]*2))

    Prof_Body = list(zip(*[iter(Prof_Body)]*2))

    Contact = list(zip(*[iter(Contact)]*2))

    Name = list(zip(*[iter(Name)]*2))

    Prof = Name + Prof_Body + Prof + Contact

    Prof = pd.DataFrame(Prof)

    Prof = Prof.transpose()

    new_header = Prof.iloc[0] #grab the first row for the header
    Prof = Prof[1:] #take the data less the header row
    Prof.columns = new_header

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    

    Exec = pd.read_html(html)[1]
    return Prof,Exec

def Historical_iv():
    driver.find_element_by_xpath('//*[@id="pairSublinksLevel2"]/li[3]').click()
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID, 'data_interval')))

    driver.find_element_by_id('data_interval').click()
    time.sleep(5)
    driver.find_element_by_xpath('//*[@id="data_interval"]/option[2]').click()
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))

    driver.find_element_by_id('widgetFieldDateRange').click()
    time.sleep(5)
    driver.find_element_by_id('startDate').clear()
    time.sleep(5)
    textarea = driver.find_element_by_id('startDate')
    time.sleep(5)
    textarea.send_keys("1/1/2014")
    time.sleep(5)

    driver.find_element_by_id('applyBtn').click()
    time.sleep(5)

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    
    html = driver.execute_script('return document.body.innerHTML;')

    T1 = pd.read_html(html)[0]

    return T1






def Ind_iv():
    driver.find_element_by_xpath('//*[@id="pairSublinksLevel2"]/li[5]').click()
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))

    df = pd.read_html(html)[0]
    return df


def News_iv():
    driver.find_element_by_xpath('//*[@id="pairSublinksLevel1"]/li[3]').click()
    driver.find_element_by_class_name('mediumTitle1')

    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    news = []
    links = []

    # for i,j,k in zip(soup.find_all('article',{'class':'js-article-item'}),soup.find_all('a',{'class':'js-external-link'}),soup.find_all('span',{'class':'articleDetails'})):

    #     span = k.text
    #     span = span.replace('\xa0-\xa0','BREAK')
    #     span = span.split('BREAK')

    #     news.append([i.text,span,str('www.investing.com'+j['href'])])

    # for i in soup.findAll('div',{'class':'mediumTitle1'}):
    #     news.append(i.text.split())

    c = 0

    for i in soup.findAll('article',{'class':'js-external-link-wrapper'}):
        link = i.get('data-link')
        l =  i.text.splitlines()
        while("" in l) :
            l.remove("")

        l.append(link)



        l[1] = l[1].split('\xa0-\xa0')
        l[1] = l[1][0] + ' ' + l[1][1]

        news.append(l)

    for i in soup.findAll('article',{'class':'js-article-item'}):
        link = i.findChildren("a" , recursive=False)
        link = re.findall("href=[\"\'](.*?)[\"\']", str(link))
        link = 'investing.com'+link[:5][0]
        l =  i.text.splitlines()
        while("" in l) :
            l.remove("")

        while(" " in l) :
            l.remove(" ")

        l.append(link)

        if('\xa0-\xa0' in i):
            l[1] = l[1].split('\xa0-\xa0')
            l[1] = l[1][0] + ' ' + l[1][1]
        

        news.append(l)

    newsfin = []
    news2 = []
    for i in news:
        if(len(i)==4):
            newsfin.append(i)
        else:
            news2.append(i)
        
    newsfin = pd.DataFrame(newsfin)

    return newsfin

def cleaner(df):
    l2 = []
    l = []

    for i in range(len(df)):
        for key,value in df.iteritems():
            l.append([value][0][i])
        l.append("")

    fin = []
    temp = []
    for i in l:
        if(i==''):
            fin.append(temp)
            temp = []
        else:
            temp.append(i)
    
    for i in fin:
        result = all(element == None for element in i)
        if(result):
            fin.remove(i)



    for i in fin:
        result = all(element == i[0] for element in i)
        if(result):
            fin.remove(i)

    fin = pd.DataFrame(fin)

    return fin

def Financial_iv():
    driver.find_element_by_xpath('//*[@id="pairSublinksLevel1"]/li[4]').click()
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')


    #Quarterly
    l1 = driver.find_element_by_xpath('//*[@id="rsdiv"]/div[1]/h3/a').get_attribute('href')
    l2 = driver.find_element_by_xpath('//*[@id="rsdiv"]/div[3]/h3/a').get_attribute('href')
    l3 = driver.find_element_by_xpath('//*[@id="rsdiv"]/div[5]/h3/a').get_attribute('href')
    l4 = driver.find_element_by_xpath('//*[@id="pairSublinksLevel2"]/li[5]/a').get_attribute('href')
    l5 = driver.find_element_by_xpath('//*[@id="pairSublinksLevel2"]/li[6]/a').get_attribute('href')
    l6 = driver.find_element_by_xpath('//*[@id="pairSublinksLevel2"]/li[7]/a').get_attribute('href')

    driver2 = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options)

    driver2.get(l1)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    html = driver2.execute_script('return document.body.innerHTML;')
    Inc_Stmt_Quart = pd.read_html(html)[1]
    Inc_Stmt_Quart = cleaner(Inc_Stmt_Quart)

    driver2.find_element_by_xpath('//*[@class="float_lang_base_1"]/a[1]').click()
    time.sleep(5)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))

    html = driver2.execute_script('return document.body.innerHTML;')
    Inc_Stmt_Ann = pd.read_html(html)[0]
    Inc_Stmt_Ann = cleaner(Inc_Stmt_Ann)



    #Balance Sheet
    driver2.get(l2)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    html = driver2.execute_script('return document.body.innerHTML;')
    Bal_Stmt_Quart = pd.read_html(html)[1]
    Bal_Stmt_Quart = cleaner(Bal_Stmt_Quart)

    driver2.find_element_by_xpath('//*[@class="float_lang_base_1"]/a[1]').click()
    time.sleep(5)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    html = driver2.execute_script('return document.body.innerHTML;')
    Bal_Stmt_Ann = pd.read_html(html)[0]
    Bal_Stmt_Ann = cleaner(Bal_Stmt_Ann)

    #CAsh Flow

    driver2.get(l3)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    html = driver2.execute_script('return document.body.innerHTML;')
    Cash_Stmt_Quart = pd.read_html(html)[1]
    Cash_Stmt_Quart = cleaner(Cash_Stmt_Quart)


    driver2.find_element_by_xpath('//*[@class="float_lang_base_1"]/a[1]').click()
    time.sleep(5)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    html = driver2.execute_script('return document.body.innerHTML;')
    Cash_Stmt_Ann = pd.read_html(html)[0]
    Cash_Stmt_Ann = cleaner(Cash_Stmt_Ann)

    

    driver2.get(l4)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    html = driver2.execute_script('return document.body.innerHTML;')
    Ratios = pd.read_html(html)[1]

    Ratios = cleaner(Ratios)
    Ratios.columns = ["Name","Company","Industry"]
    



    driver2.get(l5)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    
    
    try:
        show = driver2.find_element_by_id('showMoreDividendsHistory')
        for i in range(10):
            try:
                show.click()
                time.sleep(1)
            except:
                break
        html = driver2.execute_script('return document.body.innerHTML;')

        Div = pd.read_html(html)[0]
        Div = cleaner(Div)
    except:
        l = ["Not available"]
        Div = pd.DataFrame(l)

    



    driver2.get(l6)
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))
    try:
        clic = driver2.find_element_by_class_name('showMoreReplies')
        for i in range(10):
            try:
                clic.click()
                time.sleep(1)
            except:
                break
        html = driver2.execute_script('return document.body.innerHTML;')
        
        Er = pd.read_html(html)[0]
        Er = cleaner(Er)
    except:
        l = ["Not available"]
        Div = pd.DataFrame(l)



    driver2.close()



    return Inc_Stmt_Quart,Inc_Stmt_Ann,Bal_Stmt_Quart,Bal_Stmt_Ann,Cash_Stmt_Quart,Cash_Stmt_Ann,Ratios,Div,Er

def Technical_iv():
    driver.find_element_by_xpath('//*[@id="pairSublinksLevel1"]/li[5]').click()

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID, 'technicalstudiesSubTabs')))
    driver.find_element_by_link_text('Monthly').click()

    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'genTbl')))

    time.sleep(5)


    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')


    for i in soup.find_all("div",{"class":"newTechStudiesRight"}):
        span = i.findChildren("span")
    l=[]
    for i in span:
        l.append(i.text)

    l.remove(l[0])
    l = list(zip(*[iter(l)]*4))
    l = pd.DataFrame(l)

    

    

    T2 = pd.read_html(html)[1]

    T3 = pd.read_html(html)[2]

    T4 = pd.read_html(html)[3]

    return l,T2,T3,T4

def percdone(percent_done):
    percent_done = percent_done + 14.28
    round(percent_done, 2)
    sys.stdout.write('\r')
    sys.stdout.write("{:0.2f}% Completed".format(percent_done))
    sys.stdout.flush()
    time.sleep(0.05)
    return percent_done


def main(Company_name):
    today = date.today()
    exceldate = today.strftime("%b-%d-%Y")

    error_list = ["Cannot scrape"]

    percent_done = 0

    try:
        Summary = Summary_iv()
        percent_done = percdone(percent_done)
    except:
        Summary = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)

    try:
        Profile,Executives = Profile_iv()
        percent_done = percdone(percent_done)
    except:
        Profile,Executives = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)

    try:
        Historical_Data = Historical_iv()
        percent_done = percdone(percent_done)
    except:
        Historical_Data = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)


    try:
        Index_Component = Ind_iv()
        percent_done = percdone(percent_done)
    except:
        Index_Component = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)


    try:
        News = News_iv()
        percent_done = percdone(percent_done)
    except:
        News = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)


    try:
        Summary2,Pivot_pts,Tech_Ind,Moving_Avg = Technical_iv()
        percent_done = percdone(percent_done)
    except:
        Summary2,Pivot_pts,Tech_Ind,Moving_Avg = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)



    try:
        Inc_Stmt_Quart,Inc_Stmt_Ann,Bal_Stmt_Quart,Bal_Stmt_Ann,Cash_Stmt_Quart,Cash_Stmt_Ann,Ratios,Div,Er = Financial_iv()
        percent_done = percdone(percent_done)
    except:
        Inc_Stmt_Quart,Inc_Stmt_Ann,Bal_Stmt_Quart,Bal_Stmt_Ann,Cash_Stmt_Quart,Cash_Stmt_Ann,Ratios,Div,Er = pd.DataFrame(error_list)
        percent_done = percdone(percent_done)


    dflist= [Summary,Profile,Executives,Historical_Data,News,Index_Component,Inc_Stmt_Quart,Inc_Stmt_Ann,Bal_Stmt_Quart,Bal_Stmt_Ann,Cash_Stmt_Quart,Cash_Stmt_Ann,Ratios,Div,Er]

    for i in dflist:
        for col in i.columns[1:]:
            try:
                i[col] = i[col].str.replace(',', '').astype(float)
            except:
                i[col] = i[col]

      
    # We'll define an Excel writer object and the target file
    Excel_File_Name = str(exceldate) + '_' + Company_name + ".xlsx"
    # Excel_File_Name = os.path.join(dir_path, Excel_File_Name)
    Excelwriter = pd.ExcelWriter(Excel_File_Name,engine="xlsxwriter",engine_kwargs={'options': {'strings_to_numbers': False}})

    sheet_list = []
    #We now loop process the list of dataframes
    for df in dflist:
        sheet_list.append(retrieve_name(df)[0])
        try:
            df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],startrow=1,index=False)
        except:
            df = df.reset_index()
            df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],startrow=1,index=False)

    Summary2.to_excel(Excelwriter,sheet_name='Tech_Summary',startrow= 1,index=False)
    Pivot_pts.to_excel(Excelwriter,sheet_name='Tech_Summary',startrow= Summary2.shape[0] + 5,index=False)
    Tech_Ind.to_excel(Excelwriter,sheet_name='Tech_Summary',startrow= Pivot_pts.shape[0] + 5 + Summary2.shape[0] + 5,index=False)
    Moving_Avg.to_excel(Excelwriter,sheet_name='Tech_Summary',startrow= Tech_Ind.shape[0] + 5 + Pivot_pts.shape[0] + 5 + Summary2.shape[0] + 5,index=False)

    Excelwriter.save()



def recur():
    Company_name_list = []

    while(1):
        user_input = input("ENTER TICKER NAME (TYPE 'START' AND PRESS ENTER TO STOP READING AND START EXTRACTING): ")
        if(user_input=="START"):
            break

        if(user_input=="STOP"):
            quit()
        Company_name_list.append(user_input)

    for Company_name in Company_name_list:
        First_Page = driver.get('https://www.investing.com/search/?q='+Company_name)
        time.sleep(5)
        driver.find_element_by_xpath('//*[@class="searchSectionMain"]/div/a[1]').click()
        Current_URL = str(driver.current_url)
        main(Company_name)
        print("EXCEL FILE DOWNLOADED SUCCESSFULLY FOR --->",Company_name)
    recur()

recur()


driver.close()