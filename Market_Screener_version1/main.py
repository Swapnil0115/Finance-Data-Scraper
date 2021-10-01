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
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--allow-running-insecure-content')
chrome_options.add_argument("--disable-in-process-stack-traces")
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument("--silent")
# chrome_options.add_argument("--disables-notifications")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
# chrome_options.add_experimental_option("excludeSwitches", ["disable-popup-blocking"])
chrome_options.add_experimental_option("prefs", {"profile.default_content_setting_values.cookies": 2})

global driver
driver = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options)

def header_change(T5):
    new_header = T5.iloc[0] #grab the first row for the header
    T5 = T5[1:] #take the data less the header row
    T5.columns = new_header
    return T5

def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]

def Summary_Mrkt():
    driver2.get(str(Current_URL)+'/quotes/')
    # driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table/tbody/tr/td[3]').click()
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'tabElemNoBor')))
    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')



    T5 = pd.read_html(html)[16]

    new_header = T5.iloc[0] #grab the first row for the header
    T5 = T5[1:] #take the data less the header row
    T5.columns = new_header

    T6 = []

    for i in soup.find_all('div','fvtDiv'):
        tabtitle = i.findChildren('td',{'class':'tabTitleLeftWhite'})
        for l in tabtitle:
            tabtit = l.text
        left = i.findChildren('td',{'class':'fvtCorps2'})
        right = i.findChildren('td',{'class':'fvtCorps3'})
        temp = []
        temp.append([tabtit,''])
        for j,k in zip(left,right):
            temp.append([j.text,k.text])
        temp.append(['',''])
        T6.append(temp)


    T7 = []
    for i in soup.find_all('table','tabElemNoBor'):
        left = i.findChildren('td',{'class':'fvtCorps1'})
        temp = []
        for j in left:
            temp.append(j.text)
        temp = list(zip(*[iter(temp)]*3))
        T7.append(temp)

    T7 = [ele for ele in T7 if ele != []]

    Last_Transac = T7[0]
    Opening_hrs = T7[1]

    Last_Transac = pd.DataFrame(Last_Transac,columns=["Time","Price","Quantity"])
    Opening_hrs = pd.DataFrame(Opening_hrs,columns=["Day","Opening","Closing"])



    T6 = [ele for ele in T6 if ['Opening hours Nasdaq', ''] not in ele]

    temp2 = []
    for i in T6:
        df = pd.DataFrame(i)
        # df.drop(df.tail(1).index,inplace=True)
        temp2.append(df)

    temp2 = pd.concat(temp2)


    return T5,temp2,Last_Transac,Opening_hrs

def News_Mrkt():
    #WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'tabElemNoBor')))
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'std_txt')))

    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    try:
        next = driver2.find_element_by_xpath('//*[@class="nPageEndTab"]')
        # links = []

        T1 = pd.read_html(html)[16]

        # for i in soup.find_all('table',{'class':'tabElemNoBor'}):
        #     link = i.findChildren('td',attrs={'class':'newsColCT'})
        
        # for i in link:
        #     i2 = i.findChildren('a')
        #     i2 = re.findall("href=[\"\'](.*?)[\"\']", str(i2))
        #     links.append('www.marketscreener.com'+i2[0])

        # links = pd.DataFrame(links)

        for i in range(10):

            try:
                next.click()
            except:
                break
            time.sleep(1)
            WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'tabElemNoBor')))
            html = driver2.execute_script('return document.body.innerHTML;')
            time.sleep(1)
            
            T2 = pd.read_html(html)[16]
            # links_new = []

            # for j in soup.find_all('table',{'class':'tabElemNoBor'}):
            #     links_2 = j.findChildren('td',attrs={'class':'newsColCT'})

        
            # for j in links_2:
            #     i2 = j.findChildren('a')
            #     i2 = re.findall("href=[\"\'](.*?)[\"\']", str(i2))
            #     links_new.append('www.marketscreener.com'+i2[0])

            # links_new = pd.DataFrame(links_new)

            T1 = T1.append(T2, ignore_index = True)
            # links = links.append(links_new, ignore_index= True)


        T1.columns = ["Time","Description","Source"]
        # T1["Links"] = links

    except:
        T1 = pd.read_html(html)[16]
        T1.columns = ["Time","Description","Source"]
        # links = []

        # for i in soup.find_all('table',{'class':'tabElemNoBor'}):
        #     link = i.findChildren('td',attrs={'class':'newsColCT'})

        
        # for i in link:
        #     i2 = i.findChildren('a')
        #     i2 = re.findall("href=[\"\'](.*?)[\"\']", str(i2))
        #     links.append('www.marketscreener.com'+i2[0])

        # links = pd.DataFrame(links)
        # T1["Links"] = links

    

    return T1

def News_EXT_Mrkt():
    time.sleep(5)
    driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table[1]/tbody/tr/td[4]').click()
    time.sleep(5)
    driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table[2]/tbody/tr/td[3]').click()

    T1 = News_Mrkt()


    driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table[2]/tbody/tr/td[6]').click()
    T2 = News_Mrkt()

    return T1,T2

def Rating_Mrkt():
    
    driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table[1]/tbody/tr/td[5]').click()
    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    l = []

    for i in soup.find_all('tr'):
        td1 = i.findChildren('td',attrs={'class':'tooltip_show'})
        if(not td1):
            continue
        for j in td1:
            x1 = j.text
        indx1 = re.search("\t\t", x1)
        x1 = x1[:indx1.start()]
        td2 = i.findChildren('td',attrs={'class':'fvtCorps1'})
        td2 = re.findall("title=[\"\'](.*?)[\"\']", str(td2))
        l.append([x1,td2[0]])

    l = l[3:]
    l=pd.DataFrame(l)
    return l

def Calendar_Old():
    # years = []
    # names = []
    # values_ann = []
    # values_qtr = []

    # years_ann = []
    # years_qtr = []

    

    # #Tab1 = soup.find_all('div',{'id':'Tableau_Histo_Pub1'})

    

    # # for i in tempr:
    # #     if()

    # for i in soup.find_all('td',{'class':'bc2Y'}):
    #     years.append(i.text)

    # print(years)
    # for i in years:
    #     if('Q' in i):
    #         years_qtr.append(i)
    #     else:
    #         years_ann.append(i)

    # print(years_ann,years_qtr)
    
    # for i in soup.find_all('td',{'class':'bc2T'}):
    #     if('Announcement' in i.text):
    #         break
    #     names.append(i.text)

    # names.append('STOP')

    

    # for j in soup.find_all('td',{'class':'bc2V'}):
    #     span1 = j.findChildren('span',attrs={'class':'rtPubl'})
    #     span2 = j.findChildren('span',attrs={'class':'rtPrev'})


    #     if(len(span1)==0 or len(span2)==0):
    #         continue


    #     for k in span1:
    #         span1 = k.text
    #     for l in span2:
    #         span2 = l.text


    #     if('Released' in span1):
    #         continue

    #     span1 = span1.replace(" ","")
    #     span2 = span2.replace(" ","")

    #     if(',' in span1):
    #         span1 = span1.replace(",","")

    #     if(',' in span2):
    #         span2 = span2.replace(",","")
        
    #     span1 = int(span1)
    #     span2 = int(span2)
    #     spread = ((span1-span2)/span1)*100

    #     l1 = []
    #     l1.append(span1)
    #     l1.append(span2)
    #     l1.append(spread)

    #     values_qtr.append(l1)

    # values_qtr = list(zip(*[iter(values_ann)]*len(years_ann)))

    # ctr = 0

    # for i,j in zip(values_qtr,names):
    #     ctr+=1
    #     if(j=='STOP'):
    #         break
    #     i = list(i)
    #     i.insert(0,j)
    #     i = tuple(i)
    #     values_ann.append(i)

    # del values_qtr[:ctr]
    # print(values_qtr)
    # print(values_ann)
    pass

def cal_extr():
    action = ActionChains(driver2)

    
    try:
        scrollbar = driver2.find_element_by_xpath('//*[@id="ALNI1"]/tbody')
        action.click(on_element = scrollbar)
        time.sleep(2)
        for i in range(20):
            action.send_keys(Keys.ARROW_DOWN).perform()
            
    except:
        pass

    try:
        scrollbar = driver2.find_element_by_xpath('//*[@id="ALNI2"]/tbody')
        action.click(on_element = scrollbar)
        time.sleep(2)
        for i in range(20):
            action.send_keys(Keys.ARROW_DOWN).perform()
            
    except:
        pass

    try:
        scrollbar = driver2.find_element_by_xpath('//*[@id="ALNI3"]/tbody')
        action.click(on_element = scrollbar)
        time.sleep(2)
        for i in range(20):
            action.send_keys(Keys.ARROW_DOWN).perform()
            
    except:
        pass

    try:
        scrollbar = driver2.find_element_by_xpath('//*[@id="ALNI4"]/tbody/tr[1]/td[1]')
        action.click(on_element = scrollbar)
        time.sleep(2)
        for i in range(20):
            action.send_keys(Keys.SPACE).perform()
            
    except:
        pass


    try:
        scrollbar = driver2.find_element_by_xpath('//*[@id="ALNI5"]/tbody/tr[1]/td[1]')
        action.click(on_element = scrollbar)
        time.sleep(2)
        for i in range(20):
            action.send_keys(Keys.SPACE).perform()
            
    except:
        pass

def Calendar_Mrkt():
    driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table[1]/tbody/tr/td[6]').click()
    
    
    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    tab1 = []

    #action = ActionChains(driver2)

    cal_extr()

    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')
        

    for i,j in zip(soup.find_all('div',{'class':'content_scroll'}),soup.find_all('table',{'class':'tabTitleWhite'})):
        tr = i.findChildren('tr')
        title = j.findChildren('nobr')
        tab1.append([title[0].text,'',''])
        for j in tr:
            tds = j.findChildren('td')
            tabtemp = []
            for td in tds:
                tabtemp.append(td.text)
            
            tab1.append(tabtemp)
    

    tab1 = pd.DataFrame(tab1)

    tab1.columns = ["Date","xyz","Event Name"]
    tab1 = tab1.drop(['xyz'],axis=1)



    years = []
    years_qtr = []
    years_ann = []
    for i in soup.find_all('td',{'class':'bc2Y'}):
        years.append(i.text)

    # print(years)
    for i in years:
        if('Q' in i):
            years_qtr.append(i)
        else:
            years_ann.append(i)

    yr1 = len(years_ann)+2
    yr2 = len(years_qtr)+2

    tempr = []
    for i in soup.find_all('div',{'id':'Tableau_Histo_Pub1'}):
        k = i.findChildren('td')
        for j in k:
            if('span' in str(j) and 'bc2V' in str(j)):
                # print(j,'its span')
                spans = j.findChildren('span')
                tempr2 = []
                flag = 0

                for span in spans:
                    tempr2.append(span.text)
                    if('Released' in span.text):
                        flag = 1
                if(flag==1 or len(tempr2)==2):
                    tempr2.append('Spread')
                tempr.append(tempr2)
            else:
                # print(j)
                tempr.append(j.text)

    tempr.insert(1,'xyz')

    tempr = list(zip(*[iter(tempr)]*yr1))
    tempr = pd.DataFrame(tempr)


    Annual = tempr


    tempr = []
    for i in soup.find_all('div',{'id':'Tableau_Histo_Pub2'}):
        k = i.findChildren('td')
        for j in k:
            if('span' in str(j) and 'bc2V' in str(j)):
                # print(j,'its span')
                spans = j.findChildren('span')
                tempr2 = []
                flag = 0

                for span in spans:
                    tempr2.append(span.text)
                    if('Released' in span.text):
                        flag = 1
                if(flag==1):
                    tempr2.append('Spread')
                tempr.append(tempr2)
            else:
                # print(j)
                tempr.append(j.text)

    tempr.insert(1,'xyz')

    tempr = list(zip(*[iter(tempr)]*yr2))
    tempr = pd.DataFrame(tempr)


    Quarterly = tempr

    #driver2.close()
    Annual = header_change(Annual)
    Quarterly = header_change(Quarterly)

    return Annual,Quarterly,tab1

def Profile_Mrkt():
    driver2.get(str(Current_URL+'company/'))
    #driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[2]/tbody/tr/td/table/tbody/tr/td[7]').click()

    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'linkTabBl')))

    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')

    # desc = driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[3]/tbody/tr/td[1]/table[1]/tbody/tr[2]/td/div[1]')  
    # print(desc.text)

    # emplo = driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[3]/tbody/tr/td[1]/table[1]/tbody/tr[2]/td/div[2]')
    # print(emplo.text)


    Table1 = pd.read_html(html)[17]
    Table1 = header_change(Table1)


    Table2 = pd.read_html(html)[20]
    Table2 = header_change(Table2)


    Table3 = pd.read_html(html)[23]
    Table3 = header_change(Table3)


    Table4 = pd.read_html(html)[26]
    Table4 = header_change(Table4)


    Table5 = pd.read_html(html)[29]
    Table5 = header_change(Table5)
  

    Table6 = pd.read_html(html)[32]
    Table6 = header_change(Table6)
  

    Table7 = pd.read_html(html)[33]
    Table7 = header_change(Table7)

    # details = driver2.find_element_by_xpath('//*[@id="zbCenter"]/div/span/table[3]/tbody/tr[2]/td[1]/table[8]/tbody/tr[2]/td/div/div[1]').text.splitlines()
    # print(details)

    return Table1, Table2, Table3, Table4, Table5, Table6, Table7

def Financials_Mrkt():
    driver2.get(str(Current_URL+'financials/'))
    WebDriverWait(driver2,10).until(EC.presence_of_element_located((By.CLASS_NAME, 'highcharts-series-group')))

    html = driver2.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')
    df = []
    temp = []
    # for i in soup.find_all('tr'):
    #     BordColl = i.findChildren('table',attrs={'class':'BordCollapseYear2'})
    #     for j in BordColl:
    #         temp = []
    #         tds = j.findChildren('td')
    #         for z in tds:
    #             temp.append(z.text)
    #     df.append(temp)
    
    years = []
    yearscount = []

    for i in soup.find_all('table',{'class':'BordCollapseYear2'}):
        b2cy = i.findChildren('td',attrs={'class':'bc2Y'})
        tempyears = 0
        for k in b2cy:
            tempyears+=1
        tempyears+=1
        yearscount.append(tempyears)
        temp = []
        tds = i.findChildren('td')
        for z in tds:
            temp.append(z.text)
        df.append(temp)

    

    tempzip = []

    for i,j in zip(df,yearscount):
        tempzip.append(pd.DataFrame(list(zip(*[iter(i)]*j))))

    df1 = header_change(tempzip[0])
    df2 = header_change(tempzip[1])
    df3 = header_change(tempzip[2])
    df4 = header_change(tempzip[3])
    
    return df1,df2,df3,df4


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
        Quotes,Summary,Last_Transac,Opening_hrs = Summary_Mrkt()
    except:
        #driver2.find_element_by_xpath('//*[@id="dPubBgC"]/table/tbody/tr[1]/td/div/img').click()

        Quotes,Summary = Summary_Mrkt()

    try:
        Most_Relevant, Press_Release = News_EXT_Mrkt()
    except:
        #driver2.find_element_by_xpath('//*[@id="dPubBgC"]/table/tbody/tr[1]/td/div/img').click()
        
        Most_Relevant, Press_Release = News_EXT_Mrkt()


    try:
        Ratings = Rating_Mrkt()
    except:
        #driver2.find_element_by_xpath('//*[@id="dPubBgC"]/table/tbody/tr[1]/td/div/img').click()
        
        Ratings = Rating_Mrkt()

    try:
        Annual, Quarterly, Events = Calendar_Mrkt()
    except:
        #driver2.find_element_by_xpath('//*[@id="dPubBgC"]/table/tbody/tr[1]/td/div/img').click()
        Annual, Quarterly, Events = Calendar_Mrkt()

    try:
        Sales_per_Business, Sales_per_Region, Managers, Members_of_Board, Equities, Shareholders, Market_and_Index = Profile_Mrkt()
    except:
        #driver2.find_element_by_xpath('//*[@id="dPubBgC"]/table/tbody/tr[1]/td/div/img').click()
 
        Sales_per_Business, Sales_per_Region, Managers, Members_of_Board, Equities, Shareholders, Market_and_Index = Profile_Mrkt()

    try:
        Valuation, Annual_Data, Quarterly_Data, Balance_Sheet = Financials_Mrkt()
    except:
        #driver2.find_element_by_xpath('//*[@id="dPubBgC"]/table/tbody/tr[1]/td/div/img').click()

        Valuation, Annual_Data, Quarterly_Data, Balance_Sheet = Financials_Mrkt()
    
    

    dflist= [Quotes,Summary,Last_Transac,Opening_hrs,Most_Relevant, Press_Release, Ratings, Events,Sales_per_Business, Sales_per_Region, Managers, Members_of_Board, Equities, Shareholders, Market_and_Index,Valuation, Annual_Data, Quarterly_Data, Balance_Sheet]
    dflist2 = [Annual, Quarterly]

    # for i in dflist:
    #     for col in i.columns[1:]:
    #         try:
    #             i[col] = i[col].str.replace(',', '').astype(float)
    #         except:
    #             i[col] = i[col]

      
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

    for df in dflist2:
        sheet_list.append(retrieve_name(df)[0])
        try:
            df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],startrow=1,index=False)
        except:
            df = df.reset_index()
            df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],startrow=1,index=False)

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
        First_Page = driver.get('https://www.marketscreener.com/search/?mots='+Company_name)
        time.sleep(5)
        
        driver.find_element_by_xpath('//*[@id="ALNI0"]/tbody/tr[2]/td[3]/a').click()
        time.sleep(1)
        global Current_URL
        Current_URL = str(driver.current_url)
        print(Current_URL)
        global driver2
        driver2 = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options)
        driver2.get(Current_URL)
        main(Company_name)
        print("EXCEL FILE DOWNLOADED SUCCESSFULLY FOR --->",Company_name)
        driver2.close()

    recur()

recur()

