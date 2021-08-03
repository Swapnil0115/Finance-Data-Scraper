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


driver = webdriver.Chrome('C:\webdriver\chromedriver.exe')



#----------------------------------------------------------------------------Summary----------------------------------------------------

def Summary_Extract(Company_name):
    URL_summary = "https://finance.yahoo.com/quote/" + Company_name
    driver.get(URL_summary)
    driver.implicitly_wait(10)
    html = driver.execute_script('return document.body.innerHTML;')
    # BeautifulSoup the xml
    income_soup = BeautifulSoup(html, 'lxml')


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
    print(summary_df)
    return summary_df

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

    print(stats_df)
    return stats_df

#----------------------------------------------------------------------------Historical data part----------------------------------------------------

def Historical_Extract(Company_name):
    URL_Hist = "https://finance.yahoo.com/quote/" + Company_name + "/history"
    driver.get(URL_Hist)
    driver.implicitly_wait(10)

    # startyear = int(input('Enter starting year: '))
    # startmonth = int(input('Enter starting month: '))
    # startday = int(input('Enter starting day: '))
    startyear = 2015
    startmonth = 5
    startday = 12
    #time.sleep(2)
    startdate1 = str(startmonth)+'/'+str(startday)+'/'+str(startyear)


    # endyear = int(input('Enter ending year: '))
    # endmonth = int(input('Enter ending month: '))
    # endday = int(input('Enter ending day: '))
    endyear = 2020
    endmonth = 5
    endday = 12
    #time.sleep(2)
    enddate1 = str(endmonth)+'/'+str(endday)+'/'+str(endyear)


    html2 = driver.find_element_by_tag_name('html')
    html2.send_keys(Keys.PAGE_DOWN)
    time.sleep(2)


    Time_Period_click = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div')
    Time_Period_click.click()
    driver.find_element_by_name("startDate").send_keys(startdate1)
    driver.find_element_by_name("endDate").send_keys(enddate1)
    time.sleep(2)

    Done_button = driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/div[3]/button[1]')
    Done_button.click()
    # Max_Hist_data = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div')
    # Max_Hist_data.click()
    # Max_Hist_data_button = driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/ul[2]/li[3]')
    # Max_Hist_data_button.click()
    Apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
    Apply.click()
    time.sleep(2)


    #if 5 years data then scroll till x = 150
    yeardiff = endyear - startyear

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


    print(hist_list)

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


    print(Dividends_hist)
    print(hist_list)

    #Sort the main list Row_Wise
    hist_list_final = []


    hist_list_final = list(zip(*[iter(hist_list)]*7))

    #Make a dataframe of the sorted list
    hist_df = pd.DataFrame(hist_list_final,columns=['Date', 'Open','High','Low','Close*','Adj Close**','Volume'])
    print(hist_df)
    return hist_df

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
    
    print(profile_list)

    profile_df = pd.DataFrame(profile_list)
    profile_df = profile_df.transpose()
    profile_df.columns = ['Name','Address1','Country','Phone Number','Website','Sector','Industry','Employees','Description','Governance Score (1-10, 1 being lowest risk)']
    #print(exec_data)

    print(profile_df)

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

    exec_df = pd.DataFrame(exec_data)

    print(exec_df)
    return exec_df

#----------------------------------------------------------------------------Financials part----------------------------------------------------

def Financial_Extract(Company_name, Base_Url_Financials, Show_Type):
    urlfinancial = "https://finance.yahoo.com/quote/" + Company_name + "/" + Base_Url_Financials
    driver.get(urlfinancial)

    # if(Financial_Choice == 1):
    #     #TO click Quarterly
    #     driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div').click()

    Expand = driver.find_elements_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button')[0]
    Expand.click()
    driver.implicitly_wait(20)

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
        tuplenum = 6
    elif(Show_Type == 2):
        tuplenum = 5
    elif(Show_Type == 3):
        tuplenum = 6
    
    income_data = list(zip(*[iter(income_list)]*tuplenum))

    # print(income_data)
    # time.sleep(100)

    # # # Create a DataFrame
    income_df = pd.DataFrame(income_data)
    print(income_df)
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
    return income_df

def Financial_Extract_Quarterly(Company_name, Base_Url_Financials, Show_Type):
    urlfinancial = "https://finance.yahoo.com/quote/" + Company_name + "/" + Base_Url_Financials
    driver.get(urlfinancial)

    driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div').click()

    Expand = driver.find_elements_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button')[0]
    Expand.click()
    time.sleep(10)

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
        tuplenum = 7
    elif(Show_Type == 2):
        tuplenum = 6
    elif(Show_Type == 3):
        tuplenum = 7
    
    income_data = list(zip(*[iter(income_list)]*tuplenum))

    # print(income_data)
    # time.sleep(100)

    # # # Create a DataFrame
    income_df = pd.DataFrame(income_data)
    print(income_df)
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
    return income_df

#----------------------------------------------------------------------------Main Function---------------------------------------------------------






Company_name_list = ["IBM","AMZN","BABA","TBIO","QS","CRSR","AAPL","HOG"]

for Company_name in Company_name_list:
    Summary = Summary_Extract(Company_name)

    Statistics = Statistics_Extract(Company_name)

    Historical_Data = Historical_Extract(Company_name)

    Profile = Profile_Extract(Company_name)

    Executives = Profile_Extract2(Company_name)

    Income_Statement_Annual = Financial_Extract(Company_name, "financials",1)
    Income_Statement_Quarterly = Financial_Extract_Quarterly(Company_name, "financials",1)

    Balance_Sheet_Annual = Financial_Extract(Company_name, "balance-sheet",2)
    Balance_Sheet_Quarterly = Financial_Extract_Quarterly(Company_name, "balance-sheet",2)

    Cash_Flow_Annual = Financial_Extract(Company_name, "cash-flow",3)
    Cash_Flow_Quarterly = Financial_Extract_Quarterly(Company_name, "cash-flow",3)


    #--------------------------------------------------------------Saving all the dataframes into the excel file

    def namestr(obj, namespace):
        return [name for name in namespace if namespace[name] is obj][0]


    dflist= [Income_Statement_Annual,Balance_Sheet_Annual,Cash_Flow_Annual]
    dflist= [Profile,Executives,Summary,Statistics,Historical_Data,Income_Statement_Annual,Income_Statement_Quarterly,Balance_Sheet_Annual,Balance_Sheet_Quarterly,Cash_Flow_Annual,Cash_Flow_Quarterly]
    # We'll define an Excel writer object and the target file
    Excel_File_Name = Company_name + ".xlsx"
    Excelwriter = pd.ExcelWriter(Excel_File_Name,engine="xlsxwriter")

    sheet_list = []
    #We now loop process the list of dataframes
    for i, df in enumerate (dflist):
        sheet_list.append(namestr(df, globals()))
        df.to_excel(Excelwriter, sheet_name=namestr(df, globals()),index=False)

    # Profile.to_excel(Excelwriter,sheet_name='Result',startrow=1 , startcol=0)
    # Executives.to_excel(Excelwriter,sheet_name='Result',startrow=Profile.shape[0] + 5, startcol=0)

    #startrow=1, startcol= 1
    # header_format = Excelwriter.add_format({
    #     'bold': True,
    #     'fg_color': '#6495ED',
    #     'border': 1})

    for sheet1 in sheet_list:
        # Auto-adjust columns' width
        for column in df:
            #ExcelWriter.sheets[sheet1].write(0,column,val,header_format)
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            Excelwriter.sheets[sheet1].set_column(col_idx, col_idx, column_width)

    #And finally save the file
    Excelwriter.save()


#-----------------------------------------For Profile dfs
# writer = pd.ExcelWriter('Yahoofin.xlsx',engine='xlsxwriter')
# workbook=writer.book
# worksheet=workbook.add_worksheet('Result')
# writer.sheets['Result'] = worksheet
# worksheet.write_string(0, 0, namestr(Profile, globals()))

# Profile.to_excel(writer,sheet_name='Result',startrow=1 , startcol=0)
# worksheet.write_string(Profile.shape[0] + 4, 0, namestr(Executives, globals()))


# writer.save()


#--------------------------------------------------------------Close and Exit
driver.close()


