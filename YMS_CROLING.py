from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import datetime
from datetime import date, timedelta
import time

browser = webdriver.Chrome()
browser.maximize_window() # 창 최대화
action = ActionChains(browser)

# 1. 페이지 이동
url = 'http://ngl.logisticsmax.com/'
browser.get(url) # 페이지 이동

# 2. 로그인
id = 'jun.s'
pw = 'jun5090'

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

# Change the yard_location as PHOENIX
location = browser.find_element(By.NAME, 'yard_location')
location.send_keys('PHOENIX')
time.sleep(1)

# Change the Job_type as Import
select = browser.find_element(By.NAME,'job_type') 
select.send_keys('Import')
time.sleep(1)

# Set the Start Date
today = datetime.date.today()
yesterday = today - timedelta(days=1)
print(yesterday.strftime('%m/%d'))

startdate = browser.find_element(By.NAME,'start_date')
startdate.clear()
startdate.send_keys(yesterday.strftime('%m%d'))

end_date = browser.find_element(By.NAME,'end_date')
end_date.clear()
end_date.send_keys(yesterday.strftime('%m%d'))
time.sleep(1)

# Click the search button
browser.find_element(By.NAME, 'button2').click()

# Data Extracting
purpose_count = 0
total_count = 0
df = pd.read_html(browser.page_source)[4]    

f_name = 'kpi_import.csv'
df.to_csv(f_name,encoding='utf-8-sig', index = False, header=False)

print(df)
while True:
    pass