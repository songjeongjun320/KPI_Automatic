from csv import unregister_dialect
from re import search
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import datetime
from datetime import timedelta
import time

# Chrome webdriver automatic update
from webdriver_manager.chrome import ChromeDriverManager

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
service = Service(executable_path=ChromeDriverManager().install())
browser = webdriver.Chrome(service=service, options=options)
browser.maximize_window() # 창 최대화

# id = 
# pw = 

today = datetime.date.today()
if today.weekday() == 0:
    target_date = today - timedelta(days=3) # Monday should choose Friday
else:
    target_date = today - timedelta(days=1) # Choose yesterday

def from_YMS():
    ########## YMS ################ EXTRACT #######################################

    # 1. Go to Page
    url = 'http://ngl.logisticsmax.com/'
    browser.get(url) # 페이지 이동

    # 2. Log in on YMS

    username = browser.find_element(By.NAME, 'uid') # username search
    password = browser.find_element(By.NAME, 'pwd') # password search

    username.send_keys(id) # put username
    password.send_keys(pw) # put password

    browser.find_element(By.NAME, 'button1').click() # Click the LogIn button
    browser.find_element(By.ID, 'T3').click() # Click the Daily InOut Column

    ##############Frame switch####### IMPORTATNT#############
    iframe = browser.find_element(By.NAME, 'cf_main_sub')
    browser.switch_to.frame(iframe)
    #########################################################

    # 3. Change the yard_location as PHOENIX
    location = browser.find_element(By.NAME, 'yard_location')
    location.send_keys('PHOENIX')
    time.sleep(1)

    # 4. Change the Job_type as Import
    select = browser.find_element(By.NAME,'job_type') 
    select.send_keys('Import')
    time.sleep(1)

    # 5. Set the Start and End Date
    startdate = browser.find_element(By.NAME,'start_date')
    startdate.clear()
    startdate.send_keys(target_date.strftime('%m%d'))

    end_date = browser.find_element(By.NAME,'end_date')
    end_date.clear()
    end_date.send_keys(target_date.strftime('%m%d'))
    time.sleep(1)

    browser.find_element(By.NAME, 'button2').click() # Click the search

    # 6. Data Extracting
    purpose_count = 0
    total_count = 0
    df = pd.read_html(browser.page_source)[4]    

    # 7. Extract from YMS
    for index, row in df.iterrows(): # Count Total
        if row[7] == 'NGL':
            purpose_count += 1
        total_count+=1

    result = round(purpose_count/(total_count-1)*100,2)
    print(result)
    return result

################################################################################
################## Data extract from TMS #######################################
def from_TMS():
    browser = webdriver.Chrome(service=service)
    browser.maximize_window() # 창 최대화

    # 1. Open TMS browser
    url = 'http://nglphx.logisticsmax.com/default.asp'
    browser.get(url)

    ##############Frame switch####### IMPORTATNT#############
    iframe = browser.find_element(By.NAME, 'default_sub')
    browser.switch_to.frame(iframe)
    #########################################################

    # 2. TMS Login
    username = browser.find_element(By.NAME, 'uid') # username search
    password = browser.find_element(By.NAME, 'pwd') # password search

    username.send_keys(id) # put username
    password.send_keys(pw) # put password

    LogIn_button = browser.find_element(By.NAME,'button1') # login button search
    LogIn_button.click()

    # 3. REPORT -> Statics
    Report_button = browser.find_element(By.ID,'B3')
    Report_button.click()

    ##############Frame switch####### IMPORTATNT#############
    iframe = browser.find_element(By.NAME, 'cy_main_sub')
    browser.switch_to.frame(iframe)
    #########################################################

    Statics_button = browser.find_element(By.ID,'B10')
    Statics_button.click()

    ##############Frame switch####### IMPORTATNT#############
    iframe = browser.find_element(By.NAME, 'create_sub')
    browser.switch_to.frame(iframe)
    #########################################################

    ##############Frame switch####### IMPORTATNT#############
    iframe = browser.find_element(By.NAME, 'menu_statistics_sub')
    browser.switch_to.frame(iframe)
    #########################################################

    ## Iframe을 건너뛰어서 바로 찾을수는 없음. 순차적으로 찾을것.

    # 4. Put the date & Search

    search_date = browser.find_element(By.NAME, 'search_date')
    ####################### Always need to be cleared first ############################
    search_date.clear()
    ####################### Always need to be cleared first ############################
    tmp = target_date.strftime('%m%d')
    search_date.send_keys(tmp)
    search_button = browser.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td/fieldset/table/tbody/tr/td[1]/b/img[2]') # Click the search
    search_button.click()

    print("Sleep Start")
    time.sleep(100)
    print("Sleep End")

    # 5. Data Extracting
    ##############Frame switch####### IMPORTATNT#############
    iframe = browser.find_element(By.NAME, 'iframe4sub')
    browser.switch_to.frame(iframe)
    #########################################################
    
    pd.set_option('display.max_rows', 500) ## Show whole data    
    df = pd.read_html(browser.page_source)[3]
    
    ################## Result ###############################
    result = round(int(df.iloc[78][target_date.weekday()+2])/int(df.iloc[77][target_date.weekday()+2]) * 100, 2)
    print(result)
    return result

############################ Data Extract from OTTR #################################
def from_OTTR():
    url = 'https://drive.google.com/drive/my-drive'
    browser.get(url) # 페이지 이동

    # 1. User ID
    google_id = 'jun.s@ngltrans.net'
    google_password = 'rhkgkrwk320!'
    
    userid = browser.find_element(By.NAME,'identifier')
    userid.send_keys(google_id)

    userid_button = browser.find_element(By.XPATH,'//*[@id="identifierNext"]/div/button/span')
    userid_button.click()
    
    time.sleep(5)

    userpw = browser.find_element(By.NAME, 'Passwd')
    userpw.send_keys(google_password)

    userpw_button = browser.find_element(By.XPATH,'//*[@id="passwordNext"]/div/button/span')
    userpw_button.click()
    
    time.sleep(5)

    # 2. OTTR Search
    search_OTTR = browser.find_element(By.XPATH, '//*[@id="gs_lc50"]/input[1]')
    search_OTTR.send_keys('OTTR')
    search_OTTR.send_keys(Keys.ENTER)
    


    while True:
        pass
    
    Kpi_OTTR = 0
    
    return Kpi_OTTR


############## MAIN ################

Kpi_YMS = from_YMS()
Kpi_TMS = from_TMS()
# Kpi_OTTR = from_OTTR()

print("\n#################################################")
print("Export % : {0}".format(Kpi_TMS))
print("Chassis Flip % : {0}".format(Kpi_YMS))
print("#################################################\n")

while True:
    pass
