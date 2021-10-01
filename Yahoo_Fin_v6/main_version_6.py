from selenium import webdriver
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests
import lxml
import urllib.request as ur
import warnings
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
from requests_html import HTMLSession
from urllib.parse import urlparse
from datetime import date
from selenium.webdriver.common.action_chains import ActionChains


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

# driver = webdriver.Remote(desired_capabilities=webdriver.DesiredCapabilities.HTMLUNIT)

#chrome_options.add_argument('headless')
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--log-level=3')
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
driver = webdriver.Chrome('C:\webdriver\chromedriver.exe',chrome_options=chrome_options)
#driver.set_window_position(-10000,0)
#----------------------------------------------------------------------------Summary----------------------------------------------------

percent_done = 0

def Summary_Extract(Company_name_list,Company_name):
    URL_summary = "https://finance.yahoo.com/quote/" + Company_name
    driver.get(URL_summary)
    driver.implicitly_wait(10)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    # if("lookup" in str(driver.current_url)):
    #     driver.close()
    #     Company_name_list.remove(Company_name)
    #     print("PLEASE ENTER A VALID TICKER FOR "+str(Company_name))
    #     while(1):
    #         user_input = input("ENTER TICKER NAME (TYPE 'START' AND PRESS ENTER TO STOP READING TICKERS): ")
    #         if(user_input=="START"):
    #             break
    #         Company_name_list.append(user_input)
    #     print(Company_name_list)
    #     main_fun(Company_name_list)
    
    
    # ## Find relevant data structures for financials
    summary_list = []

    # Find all HTML data structures that are divs
    for div in income_soup.find_all('td'):
        # Get the contents and titles
        summary_list.append(div.text)

    summary_list = list(filter(None, summary_list))

    summary_list_final = []
    i_summ = 0

    while(i_summ!=len(summary_list)):
        key = summary_list[i_summ]
        val = summary_list[i_summ+1]
        summary_list_final.append([key,val])
        i_summ = i_summ+2

    summary_df = pd.DataFrame(summary_list_final,columns=['Summary', 'Value'])

    for col in summary_df.columns[1:]:                  # UPDATE ONLY NUMERIC COLS 
        try:
            summary_df[col] = summary_df[col].str.replace(',', '').astype(float)
            #print(summary_df[col])
        except:
            summary_df.loc[summary_df[col] == '-', col] = np.nan    # REPLACE HYPHEN WITH NaNs
    
    #print(percent_done+=5.5)
    return summary_df

def News_Extract(Company_name):
    session = HTMLSession()
    r = session.get("https://finance.yahoo.com/quote/" + Company_name + "/news")
    r.html.render(scrolldown = 5000)

    news = r.html.find('.js-stream-content',first=False)
    
    dates = []
    h3s = []
    ps = []
    urls = []
    news_list = []

    for i in news:
        if('simple-list-item' in str(i)):
            break
        dates.append(i.find('span')[1].text)
        h3s.append(i.find('h3')[0].text)
        ps.append(i.find('p')[0].text)
        temp = i.links
        for i in temp:
            tempurl = urlparse("https://finance.yahoo.com"+str(i))
            urls.append(tempurl.geturl())
        
    for i in range(len(h3s)):
        news_list.append([h3s[i],urls[i],ps[i],dates[i]])

    #print(h3s,ps,urls,dates)

    #print(news_list)
    news_df = pd.DataFrame(news_list,columns = ["Article Heading","Article Link","Article Description","Article Date"])
    #print(news_df)
    return news_df

def Press_Extract(Company_name):
    session = HTMLSession()
    r = session.get("https://finance.yahoo.com/quote/" + Company_name + "/press-releases")
    r.html.render(scrolldown = 5000)

    news = r.html.find('.js-stream-content',first=False)
    
    dates = []
    h3s = []
    ps = []
    urls = []
    news_list = []

    for i in news:
        if('simple-list-item' in str(i)):
            break
        dates.append(i.find('span')[1].text)
        h3s.append(i.find('h3')[0].text)
        ps.append(i.find('p')[0].text)
        temp = i.links
        for i in temp:
            tempurl = urlparse("https://finance.yahoo.com"+str(i))
            urls.append(tempurl.geturl())
        
    for i in range(len(h3s)):
        news_list.append([h3s[i],urls[i],ps[i],dates[i]])

    #print(h3s,ps,urls,dates)

    #print(news_list)
    press_df = pd.DataFrame(news_list,columns = ["Article Heading","Article Link","Article Description","Article Date"])
    #print(press_df)
    return press_df

    
#----------------------------------------------------------------------------Statistics part----------------------------------------------------

def Statistics_Extract(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/key-statistics"
    driver.get(URL_stat)
    driver.implicitly_wait(10)

    html = driver.execute_script('return document.body.innerHTML;')


    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    # ## Find relevant data structures for financials
    stats_list = []

    # Find all HTML data structures that are divs
    for div in income_soup.find_all('td'):
        # if(div.find('span')):
        #     #print(div.find('span').text)
        #     td_list.append([div.find('span').text,div.string])

        stats_list.append(div.text)
        # print(div.text)
        # Prevent duplicate titles
        # if not div.string == div.get('title'):
        #     td_list.append(div.get('title'))


    #td_list = list(filter(None, td_list))
    #print(td_list)

    stats_list_final = []
    i = 0

    while(i!=len(stats_list)):
        key = stats_list[i]
        val = stats_list[i+1]
        stats_list_final.append([key,val])
        i = i+2

    stats_df = pd.DataFrame(stats_list_final,columns=['Statistics', 'Value'])

    #print(stats_df)
    return stats_df

#----------------------------------------------------------------------------Historical data part----------------------------------------------------

def Historical_Extract(Company_name):
    URL_Hist = "https://finance.yahoo.com/quote/" + Company_name + "/history"
    driver.get(URL_Hist)

    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    time.sleep(3)
    html2 = driver.find_element_by_tag_name('html')
    html2.send_keys(Keys.PAGE_DOWN)

    WebDriverWait(driver, 10)
    time.sleep(3)

    Time_Period_click = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div')
    action = ActionChains(driver)
    action.click(on_element = Time_Period_click)
    action.perform()

    Max_Data = driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/ul[2]/li[3]/button')
    action = ActionChains(driver)
    action.click(on_element = Max_Data)
    action.perform()

    Apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
    action = ActionChains(driver)
    action.click(on_element = Apply)
    action.perform()


    # try:
    #     wait = WebDriverWait(driver, 10)
    #     EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div'))
    #     Time_Period_click = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div')
    # except:
    #     try:
    #         time.sleep(5)
    #         wait = WebDriverWait(driver, 10)
    #         EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div'))
    #         html2.send_keys(Keys.PAGE_DOWN)
    #         Time_Period_click = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div')
    #         Time_Period_click.click()
    #     except:
    #         time.sleep(10)
    #         wait = WebDriverWait(driver, 10)
    #         EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div'))
    #         html2.send_keys(Keys.PAGE_DOWN)
    #         Time_Period_click = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div')
    #         Time_Period_click.click()

    # time.sleep(1)
    
    # try:
    #     wait = WebDriverWait(driver, 10)
    #     EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[2]/section/div[1]/div[1]/div[1]/div/div/div/div/div/ul[2]/li[4]'))
    #     Max_Data = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[2]/section/div[1]/div[1]/div[1]/div/div/div/div/div/ul[2]/li[4]')
    #     Max_Data.click()
    # except:
    #     try:
    #         time.sleep(5)
    #         EC.element_to_be_clickable((By.XPATH, '//*[@id="dropdown-menu"]/div/ul[2]/li[4]'))
    #         Max_Data = driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/ul[2]/li[4]')
    #         Max_Data.click()
    #     except:
    #         try:
    #             time.sleep(10)
    #             EC.element_to_be_clickable((By.XPATH, '//*[@id="dropdown-menu"]/div/ul[2]/li[4]/button'))
    #             Max_Data = driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/ul[2]/li[4]/button')
    #             Max_Data.click()
    #         except:
    #             Historical_Extract(Company_name)

    time.sleep(1)

    startyear_str = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div/div/div/span').text
    startyear = startyear_str[8:12]

    # try:
    #     time.sleep(2)
    #     EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button'))
    #     Apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
    #     Apply.click()
        
    # except:
    #     try:
    #         time.sleep(5)
    #         EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button'))
    #         Apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
    #         Apply.click()
    #     except:
    #         try:
    #             time.sleep(10)
    #             EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button'))
    #             Apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
    #             Apply.click()
    #         except:
    #             time.sleep(15)
    #             EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button'))
    #             Apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
    #             Apply.click()

    time.sleep(3)
    today = date.today()
    endyear = today.strftime("%Y")
    


    #if 5 years data then scroll till x = 150
    yeardiff = int(endyear) - int(startyear)

    if(yeardiff<=5):
        yeardiffscroll = 150
    elif(yeardiff>5 and yeardiff<=10):
        yeardiffscroll = 320
    elif(yeardiff>10 and yeardiff<=15):
        yeardiffscroll = 470
    elif(yeardiff>15 and yeardiff<=20):
        yeardiffscroll = 620
    elif(yeardiff>20 and yeardiff<=25):
        yeardiffscroll = 770
    elif(yeardiff>25 and yeardiff<=30):
        yeardiffscroll = 920
    elif(yeardiff>30 and yeardiff<=35):
        yeardiffscroll = 1070
    elif(yeardiff>36):
        yeardiffscroll = 2000

    x = 0
    while(x!=yeardiffscroll):
        html2.send_keys(Keys.PAGE_DOWN)
        x = x + 1

    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    # Find all HTML data structures that are tds
    hist_list = []
    for div in income_soup.find_all('td'):
        hist_list.append(div.text)


    #print(hist_list)

    #Move all the dividends info to Dividends List and delete all the useless info in the end of the hist_list
    Dividends_hist = []
    Stock_Split = []
    #Remove dividend rows
    for i_hist,val in enumerate(hist_list):
        if("Dividend" in val):
            Dividends_hist.append(hist_list[i_hist-1:i_hist+1])
            del hist_list[i_hist-1:i_hist+1]
            # print("Success")
        elif("Stock Split" in val):
            Stock_Split.append(hist_list[i_hist-1:i_hist+1])
            del hist_list[i_hist-1:i_hist+1]
        elif("*Close price adjusted for splits" in val):
            del hist_list[i_hist:]


    #print(Dividends_hist)
    #print(hist_list)

    #Sort the main list Row_Wise
    hist_list_final = []


    hist_list_final = list(zip(*[iter(hist_list)]*7))

    #Make a dataframe of the sorted list
    hist_df = pd.DataFrame(hist_list_final,columns=['Date', 'Open','High','Low','Close*','Adj Close**','Volume'])
    for col in hist_df.columns[1:]:                  # UPDATE ONLY NUMERIC COLS 
        try:
            hist_df[col] = hist_df[col].str.replace(',', '').astype(float)
            #print(hist_df[col])
        except:
            hist_df.loc[hist_df[col] == '-', col] = np.nan    # REPLACE HYPHEN WITH NaNs
            
    Dividends_hist_df = pd.DataFrame(Dividends_hist,columns = ["Date","Dividends"])
    Stock_Split_df = pd.DataFrame(Stock_Split,columns=['Date', 'Splits'])
    #print(hist_df)
    return hist_df,Dividends_hist_df,Stock_Split_df

#----------------------------------------------------------------------------Profile-----------------------------------------------------------------

def Profile_Extract(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/profile"
    driver.get(URL_stat)
    driver.implicitly_wait(10)
    
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    #profile_list = []
    prof = driver.find_element_by_xpath('//*[@id="Col1-0-Profile-Proxy"]/section/div[1]/div')
    profile_list = prof.text.splitlines()

    Descrip = driver.find_element_by_xpath('//*[@id="Col1-0-Profile-Proxy"]/section/section[2]/p')
    GovtScore = driver.find_element_by_xpath('//*[@id="Col1-0-Profile-Proxy"]/section/section[3]/div')
    profile_list.append(Descrip.text)
    profile_list.append(GovtScore.text)
    
    #print(profile_list)

    profile_df = pd.DataFrame(profile_list)
    profile_df = profile_df.transpose()
    #profile_df.columns = ['Name','Address1','Country','Phone Number','Website','Sector','Industry','Employees','Description','Governance Score (1-10, 1 being lowest risk)']
    #print(exec_data)

    #print(profile_df)

    return profile_df

def Profile_Extract2(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/profile"
    driver.get(URL_stat)
    driver.implicitly_wait(10)
    
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    exec_list = []

    for div in income_soup.find_all('td'):
        exec_list.append(div.text)

    exec_data = list(zip(*[iter(exec_list)]*5))

    exec_df = pd.DataFrame(exec_data,columns = ["Name","Title","Pay","Exercised","Year Born"])

    #print(exec_df)
    return exec_df

#----------------------------------------------------------------------------Financials part----------------------------------------------------

def Financial_Extract(Company_name, Base_Url_Financials, Show_Type):
    urlfinancial = "https://finance.yahoo.com/quote/" + Company_name + "/" + Base_Url_Financials
    driver.get(urlfinancial)

    # if(Financial_Choice == 1):
    #     #TO click Quarterly
    #     driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div').click()
    time.sleep(2)
    Expand = driver.find_elements_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button')[0]
    action = ActionChains(driver)
    action.click(on_element = Expand)
    action.perform()

    driver.implicitly_wait(20)

    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    time.sleep(2)

    # ## Find relevant data structures for financials
    div_list = []

    flag_fin = 0

    # Find all HTML data structures that are divs
    for div in income_soup.find_all('div'):
        # Get the contents and titles

        if(div.text == "Breakdown"):
            flag_fin = 1

        if(flag_fin == 1):
            div_list.append(div.string)

        # Prevent duplicate titles
        if not div.string == div.get('title'):
            div_list.append(div.get('title'))


    # Filter out 'empty' elements
    div_list = list(filter(None, div_list))

    try:
        tuple_num_index = div_list.index("Total Revenue")
    except:
        try:
            tuple_num_index = div_list.index("Total Assets")
        except:
            tuple_num_index = div_list.index("Operating Cash Flow")


    # Filter out functions
    div_list = [incl for incl in div_list if not incl.startswith('(function')]

    

    # # Sublist the relevant financial information
    income_list = div_list



    # # # Insert "Breakdown" to the beginning of the list to give it the proper stucture
    income_list.insert(0, 'Breakdown')

    for i,val in enumerate(income_list):
        if(val == "Advertise with us"):
            del income_list[i-1:]
            break


    # # # ## Create a DataFrame of the financial data
    # # # Store the financial items as a list of tuples

    if(Show_Type == 1):
        tuplenum = tuple_num_index+1
    elif(Show_Type == 2):
        tuplenum = tuple_num_index+1
    elif(Show_Type == 3):
        tuplenum = tuple_num_index+1
    
    income_data = list(zip(*[iter(income_list)]*tuplenum))

    # print(income_data)
    # time.sleep(100)

    # # # Create a DataFrame
    income_df = pd.DataFrame(income_data)
    #print(income_df)
    time.sleep(5)
    # # Make the top row the headers
    # headers = income_df.iloc[0]
    # income_df = income_df[1:]
    # income_df.columns = headers
    # income_df.set_index('Breakdown', inplace=True, drop=True)

    new_header = income_df.iloc[0] #grab the first row for the header
    income_df = income_df[1:] #take the data less the header row
    income_df.columns = new_header #set the header row as the df header
    # warnings.warn('Amounts are in thousands.')
    income_df = income_df.iloc[:, ::-1]

    # shift column 'C' to first position
    first_column = income_df.pop('Breakdown')
    
    # insert column using insert(position,column_name,first_column) function
    income_df.insert(0, 'Breakdown', first_column)
    for col in income_df.columns[1:]:                  # UPDATE ONLY NUMERIC COLS 
        try:
            income_df[col] = income_df[col].str.replace(',', '').astype(float)
            #print(income_df[col])
        except:
            income_df.loc[income_df[col] == '-', col] = np.nan    # REPLACE HYPHEN WITH NaNs

    return income_df

def Financial_Extract_Quarterly(Company_name, Base_Url_Financials, Show_Type):
    urlfinancial = "https://finance.yahoo.com/quote/" + Company_name + "/" + Base_Url_Financials
    driver.get(urlfinancial)

    time.sleep(2)

    WebDriverWait(driver, 10)
    EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div'))
    Expand1 = driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div')
    action = ActionChains(driver)
    action.click(on_element = Expand1)
    action.perform()
    time.sleep(2)

    EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button'))
    Expand = driver.find_elements_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button')[0]
    action = ActionChains(driver)
    action.click(on_element = Expand)
    action.perform()
    time.sleep(3)

    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')


    # ## Find relevant data structures for financials
    div_list = []

    flag_fin = 0

    # Find all HTML data structures that are divs
    for div in income_soup.find_all('div'):
        # Get the contents and titles
        if(div.text == "Breakdown"):
            flag_fin = 1

        if(flag_fin == 1):
            div_list.append(div.string)

        # Prevent duplicate titles
        if not div.string == div.get('title'):
            div_list.append(div.get('title'))


    # Filter out 'empty' elements
    div_list = list(filter(None, div_list))

    try:
        tuple_num_index = div_list.index("Total Revenue")
    except:
        try:
            tuple_num_index = div_list.index("Total Assets")
        except:
            tuple_num_index = div_list.index("Operating Cash Flow")
            

    # Filter out functions
    div_list = [incl for incl in div_list if not incl.startswith('(function')]


    # # Sublist the relevant financial information
    income_list = div_list



    # # # Insert "Breakdown" to the beginning of the list to give it the proper stucture
    income_list.insert(0, 'Breakdown')

    for i,val in enumerate(income_list):
        if(val == "Advertise with us"):
            del income_list[i-1:]
            break


    # # # ## Create a DataFrame of the financial data
    # # # Store the financial items as a list of tuples

    if(Show_Type == 1):
        tuplenum = tuple_num_index+1
    elif(Show_Type == 2):
        tuplenum = tuple_num_index+1
    elif(Show_Type == 3):
        tuplenum = tuple_num_index+1
    
    income_data = list(zip(*[iter(income_list)]*tuplenum))

    # print(income_data)
    # time.sleep(100)

    # # # Create a DataFrame
    income_df = pd.DataFrame(income_data)
    #print(income_df)
    time.sleep(5)
    # # Make the top row the headers
    # headers = income_df.iloc[0]
    # income_df = income_df[1:]
    # income_df.columns = headers
    # income_df.set_index('Breakdown', inplace=True, drop=True)

    new_header = income_df.iloc[0] #grab the first row for the header
    income_df = income_df[1:] #take the data less the header row
    income_df.columns = new_header #set the header row as the df header
    # warnings.warn('Amounts are in thousands.')
    income_df = income_df.iloc[:, ::-1]

    # shift column 'C' to first position
    first_column = income_df.pop('Breakdown')
    
    # insert column using insert(position,column_name,first_column) function
    income_df.insert(0, 'Breakdown', first_column)

    for col in income_df.columns[1:]:                  # UPDATE ONLY NUMERIC COLS 
        try:
            income_df[col] = income_df[col].str.replace(',', '').astype(float)
            #print(income_df[col])
        except:
            income_df.loc[income_df[col] == '-', col] = np.nan    # REPLACE HYPHEN WITH NaNs

    return income_df

#-----------------------------------------------------------------------------Analysis

def Analysis_Extract(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/analysis"
    driver.get(URL_stat)
    driver.implicitly_wait(10)

    html = driver.execute_script('return document.body.innerHTML;')


    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    # ## Find relevant data structures for financials
    heading_list = []

    
    
    for i in income_soup.find_all('th'):
        heading_list.append(i.text)

    df1_heading = heading_list[0:5]
    df2_heading = heading_list[5:10]
    df3_heading = heading_list[10:15]
    df4_heading = heading_list[15:20]
    df5_heading = heading_list[20:25]
    df6_heading = heading_list[25:30]

    tds = []
    
    for i in income_soup.find_all('td'):
        tds.append(i.text)
    
    df1 = tds[0:25]
    df2 = tds[25:55]
    df3 = tds[55:75]
    df4 = tds[75:100]
    df5 = tds[100:120]
    df6 = tds[120:150]

    # df1 = df1_heading+df1
    # df2 = df2_heading+df2
    # df3 = df3_heading+df3
    # df4 = df4_heading+df4
    # df5 = df5_heading+df5
    # df6 = df6_heading+df6

    Earnings_DF = np.array(df1)
    Earnings_DF = np.reshape(Earnings_DF, (5,5))
    Earnings_DF = pd.DataFrame(Earnings_DF, columns=df1_heading)

    Rev_DF = np.array(df2)
    Rev_DF = np.reshape(Rev_DF, (6,5))
    Rev_DF = pd.DataFrame(Rev_DF, columns=df2_heading)

    Earning_hist_DF = np.array(df3)
    Earning_hist_DF = np.reshape(Earning_hist_DF, (4,5))
    Earning_hist_DF = pd.DataFrame(Earning_hist_DF, columns=df3_heading)

    EPS_DF = np.array(df4)
    EPS_DF = np.reshape(EPS_DF, (5,5))
    EPS_DF= pd.DataFrame(EPS_DF, columns=df4_heading)

    EPS_Rev_DF = np.array(df5)
    EPS_Rev_DF = np.reshape(EPS_Rev_DF, (4,5))
    EPS_Rev_DF = pd.DataFrame(EPS_Rev_DF, columns=df5_heading)

    Growth_DF = np.array(df6)
    Growth_DF = np.reshape(Growth_DF, (6,5))
    Growth_DF = pd.DataFrame(Growth_DF, columns=df6_heading)

    return Earnings_DF,Rev_DF,Earning_hist_DF,EPS_DF,EPS_Rev_DF,Growth_DF

def Holders_Extract(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/holders"
    driver.get(URL_stat)
    driver.implicitly_wait(10)

    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')

    dfs = pd.read_html(html)[0]
    dfs1 = pd.read_html(html)[1]
    dfs2 = pd.read_html(html)[2]

    return dfs,dfs1,dfs2


def Insider_Roster_Extract(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/insider-roster"
    driver.get(URL_stat)
    driver.implicitly_wait(10)
    time.sleep(10)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    #dfs = pd.read_html(html)[0]    

    td_list = []

    for i in income_soup.find_all("td",{"class":"Ta(end)"}):
        td_list.append(i.text)

    indiv_list = []
    design_list = []

    for i in income_soup.find_all("td",{"class":"Ta(start)"}):
        indiv_list.append(i.find("a").text)
        design_list.append(i.text.replace((i.find("a").text),''))

    # print(td_list)
    # print(indiv_list)
    # print(design_list)

    td_list = list(zip(*[iter(td_list)]*3))
    td_df = pd.DataFrame(td_list,columns = ["Most Recent Transaction","Date","Shares Owned as of Transaction Date"])

    td_df.insert(0,"Individual or Entity",indiv_list)
    td_df.insert(1,"Designation",design_list)

    return td_df

def Insider_Transactions_Extract1(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/insider-transactions"
    driver.get(URL_stat)
    driver.implicitly_wait(10)

    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')
    dfs = pd.read_html(html)[0]
    dfs1 = pd.read_html(html)[1]

    #print(dfs,dfs1,dfs2)
    return dfs,dfs1

def Insider_Transactions_Extract2(Company_name):
    URL_stat = "https://finance.yahoo.com/quote/" + Company_name + "/insider-transactions"
    driver.get(URL_stat)
    driver.implicitly_wait(10)
    time.sleep(5)

    temp = driver.find_element_by_xpath('//*[@id="Col1-1-Holders-Proxy"]/section/div[2]/div[4]/table/tbody').text.splitlines()
    temp = list(zip(*[iter(temp)]*3))

    td_list = []
    temp2 = []
    for i in range(1,len(temp)+1):
        for j in range(1,7):
            temp2.append(driver.find_element_by_xpath('//*[@id="Col1-1-Holders-Proxy"]/section/div[2]/div[4]/table/tbody/tr['+str(i)+']/td['+str(j)+']').text)

    for i in temp2:
        if("\n" in i):
            ind = i.find("\n")
            ele1 = i[:ind]
            ele2 = i[ind+1:]
            td_list.append(ele1)
            td_list.append(ele2)
        else:
            td_list.append(i)

    td_list = list(zip(*[iter(td_list)]*7))

    dfs2 = pd.DataFrame(td_list,columns = ["Insider","Designation","Transaction","Type","Value","Date","Shares"])
    return dfs2

def Error_Extract():
    error_list = ["Cannot scrape"]

    df1 = pd.DataFrame(error_list)
    df2 = pd.DataFrame(error_list)
    df3 = pd.DataFrame(error_list)
    df4 = pd.DataFrame(error_list)
    df5 = pd.DataFrame(error_list)
    df6 = pd.DataFrame(error_list)

    return df1,df2,df3,df4,df5,df6
    
def Error_Extract2():
    error_list = ["Cannot scrape"]

    df1 = pd.DataFrame(error_list)
    df2 = pd.DataFrame(error_list)
    df3 = pd.DataFrame(error_list)

    return df1,df2,df3

def Error_Extract3():
    error_list = ["Cannot scrape"]

    df1 = pd.DataFrame(error_list)
    df2 = pd.DataFrame(error_list)

    return df1,df2


#----------------------------------------------------------------------------Main Function---------------------------------------------------------

# def namestr(obj, namespace):
#     return [name for name in namespace if namespace[name] is obj][0]

def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]


#Error in company "BABA" in Profile
def main_fun(Company_name_list,wrong_tickers):
    today = date.today()
    exceldate = today.strftime("%b-%d-%Y")
    error_list = ["Cannot scrape"]
    
    lossy_flag = 0
    lossy_tickers = []
    for Company_name in Company_name_list:
        percent_done = 0
        try:
            Summary = Summary_Extract(Company_name_list,Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Summary = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            News = News_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            News = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Press = Press_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Press = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Statistics = Statistics_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Statistics = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Historical_Data,Dividends,Splits = Historical_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            try:
                Historical_Data,Dividends,Splits = Historical_Extract(Company_name)
                percent_done+=5.5
                print(str(percent_done)+"% Completed")
            except:
                Historical_Data,Dividends,Splits = Error_Extract2
                percent_done+=5.5
                print(str(percent_done)+"% Completed")

        try:
            Profile = Profile_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Profile = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Executives = Profile_Extract2(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Executives = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Income_Statement_Annual = Financial_Extract(Company_name, "financials",1)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Income_Statement_Annual = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Income_Statement_Quarterly = Financial_Extract_Quarterly(Company_name, "financials",1)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Income_Statement_Quarterly = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Balance_Sheet_Annual = Financial_Extract(Company_name, "balance-sheet",2)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Balance_Sheet_Annual = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Balance_Sheet_Quarterly = Financial_Extract_Quarterly(Company_name, "balance-sheet",2)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Balance_Sheet_Quarterly = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Cash_Flow_Annual = Financial_Extract(Company_name, "cash-flow",3)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Cash_Flow_Annual = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Cash_Flow_Quarterly = Financial_Extract_Quarterly(Company_name, "cash-flow",3)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Cash_Flow_Quarterly = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Earnings_Estimate,Revenue_Estimate,Earnings_History_DF,EPS_Trend,EPS_Revisions,Growth_Estimates = Analysis_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Earnings_Estimate,Revenue_Estimate,Earnings_History_DF,EPS_Trend,EPS_Revisions,Growth_Estimates = Error_Extract()
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        
        try:
            Major_Holders,Top_Institutional_Holders2,Top_Mutual_Fund_Holders = Holders_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Major_Holders,Top_Institutional_Holders2,Top_Mutual_Fund_Holders = Error_Extract2()
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        
        try:
            Insider_Roster = Insider_Roster_Extract(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Insider_Roster = pd.DataFrame(error_list)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Ins_Transac_6_mo,Net_Institutional_Transac = Insider_Transactions_Extract1(Company_name)
            percent_done+=5.5
            print(str(percent_done)+"% Completed")
        except:
            Ins_Transac_6_mo,Net_Institutional_Transac = Error_Extract3()
            percent_done+=5.5
            print(str(percent_done)+"% Completed")

        try:
            Insider_Transac_2_yr = Insider_Transactions_Extract2(Company_name)
            percent_done = 100
            print(str(percent_done)+"% Completed")
        except:
            Insider_Transac_2_yr = pd.DataFrame(error_list)
            percent_done = 100
            print(str(percent_done)+"% Completed")

        # if(lossy_flag==1):
        #     lossy_tickers.append(Company_name)
        #--------------------------------------------------------------Saving all the dataframes into the excel file

        #dflist= [Income_Statement_Annual,Balance_Sheet_Annual,Cash_Flow_Annual]
        dflist= [Profile,News,Press,Executives,Summary,Statistics,Historical_Data,Dividends,Splits,Income_Statement_Annual,Income_Statement_Quarterly,Balance_Sheet_Annual,Balance_Sheet_Quarterly,Cash_Flow_Annual,Cash_Flow_Quarterly,Earnings_Estimate,Revenue_Estimate,Earnings_History_DF,EPS_Trend,EPS_Revisions,Growth_Estimates,Major_Holders,Top_Institutional_Holders2,Top_Mutual_Fund_Holders,Insider_Roster,Ins_Transac_6_mo,Net_Institutional_Transac,Insider_Transac_2_yr]
        for i in dflist:
            for col in i.columns[1:]:
                try:
                    i[col] = i[col].str.replace(',', '').astype(float)
                except:
                    i[col] = i[col]
            
        # We'll define an Excel writer object and the target file
        Excel_File_Name = str(exceldate) + '_' + Company_name + ".xlsx"
        Excelwriter = pd.ExcelWriter(Excel_File_Name,engine="xlsxwriter",engine_kwargs={'options': {'strings_to_numbers': False}})

        sheet_list = []
        #We now loop process the list of dataframes
        for df in dflist:
            sheet_list.append(retrieve_name(df)[0])
            df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],index=False)

        # Profile.to_excel(Excelwriter,sheet_name='Result',startrow=1 , startcol=0)
        # Executives.to_excel(Excelwriter,sheet_name='Result',startrow=Profile.shape[0] + 5, startcol=0)

        for sheet1 in sheet_list:
            # Auto-adjust columns' width
            try:
                for column in df:
                    try:
                        #ExcelWriter.sheets[sheet1].write(0,column,val,header_format)
                        column_width = max(df[column].astype(str).map(len).max(), len(column))
                        col_idx = df.columns.get_loc(column)
                        Excelwriter.sheets[sheet1].set_column(col_idx, col_idx, column_width)
                    except:
                        continue
            except:
                continue

        #And finally save the file
        Excelwriter.save()
        
        print("EXCEL FILE DOWNLOADED SUCCESSFULLY FOR --->",Company_name)
    
    print("DOWNLOADED TICKERS FOR: ",Company_name_list)

    if(len(wrong_tickers)>0):
        print("PLEASE ENTER CORRECT TICKER NAMES FOR THE FOLLOWING TICKER(s)--->",wrong_tickers)

    # if(len(lossy_tickers)>0):
    #     print("SOME OF THE DATA WAS LOST WHILE SCRAPING THESE TICKER(s)--->",lossy_tickers)

    
    Company_name_list = []
    while(1):
        user_input = input("ENTER TICKER NAME (TYPE 'START' AND PRESS ENTER TO STOP READING AND START EXTRACTING): ")
        if(user_input=="START"):
            break
        Company_name_list.append(user_input)

    wrong_tickers = []
    for i in Company_name_list:
        test_URL = "https://finance.yahoo.com/quote/" + i
        driver.get(test_URL)
        if("lookup" in str(driver.current_url)):
            wrong_tickers.append(i)
            del i
    

    main_fun(Company_name_list,wrong_tickers)



def test(Company_name):
    df1 = News_Extract(Company_name)
    df2 = Press_Extract(Company_name)

    dflist= [df1,df2]
    # for col in Income_Statement_Annual.columns[1:]:
    #     try:
    #         Income_Statement_Annual[col] = Income_Statement_Annual[col].str.replace(',', '').astype(float)
    #     except:
    #         Income_Statement_Annual[col] = Income_Statement_Annual[col]

    Excel_File_Name = Company_name + ".xlsx"
    Excelwriter = pd.ExcelWriter(Excel_File_Name,engine="xlsxwriter",engine_kwargs={'options': {'strings_to_numbers': False}})

    for df in dflist:
        df.to_excel(Excelwriter, sheet_name=retrieve_name(df)[0],index=False)
    Excelwriter.save()

#---------------------------------------------------------------------------CALL FUNCTIONS---------------------------------------------------

# def correct_url_checker():
#     user_input_fun = input("ENTER TICKER NAME (TYPE 'START' AND PRESS ENTER TO STOP READING TICKERS): ")
#     if(user_input_fun=="START"):
#         return "START"
#     print("PROCESSING....")
#     test_URL = "https://finance.yahoo.com/quote/" + user_input_fun
#     driver.get(test_URL)
#     if("lookup" in str(driver.current_url)):
#         print("PLEASE ENTER A VALID TICKER FOR "+str(user_input_fun))
#         return correct_url_checker()
#     else:
#         return user_input_fun

# Company_name_list = []
# while(1):
#     user_input = correct_url_checker()
#     if(user_input=="START"):
#         break
#     Company_name_list.append(user_input)

# main_fun(Company_name_list)



Company_name_list = []
while(1):
    user_input = input("ENTER TICKER NAME (TYPE 'START' AND PRESS ENTER TO STOP READING AND START EXTRACTING): ")
    if(user_input=="START"):
        break
    Company_name_list.append(user_input)

wrong_tickers = []
for i in Company_name_list:
    test_URL = "https://finance.yahoo.com/quote/" + i
    driver.get(test_URL)
    if("lookup" in str(driver.current_url)):
        wrong_tickers.append(i)
        Company_name_list.remove(i)

#print(Company_name_list)
if(len(wrong_tickers)>0):
    print("INCORRECT TICKERS, ENTER PROPERLY IN THE NEXT RUN-->",wrong_tickers)

main_fun(Company_name_list,wrong_tickers)









#test("IBM")


#--------------------------------------------------------------Close and Exit
driver.close()


