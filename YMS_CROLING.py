from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import datetime
from datetime import date, timedelta
import time

browser = webdriver.Chrome()
browser.maximize_window() # 창 최대화

########## YMS ################ EXTRACT #######################################

# 1. Go to Page
url = 'http://ngl.logisticsmax.com/'
browser.get(url) # 페이지 이동

# 2. Log in on YMS
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

# 3. Change the yard_location as PHOENIX
location = browser.find_element(By.NAME, 'yard_location')
location.send_keys('PHOENIX')
time.sleep(1)

# 4. Change the Job_type as Import
select = browser.find_element(By.NAME,'job_type') 
select.send_keys('Import')
time.sleep(1)

# 5. Set the Start and End Date
today = datetime.date.today()
yesterday = today - timedelta(days=2)

startdate = browser.find_element(By.NAME,'start_date')
startdate.clear()
startdate.send_keys(yesterday.strftime('%m%d'))

end_date = browser.find_element(By.NAME,'end_date')
end_date.clear()
end_date.send_keys(yesterday.strftime('%m%d'))
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

Kpi_Yms = round(purpose_count/(total_count-1)*100,2)
print(Kpi_Yms) ### KPI From YMS

################################################################################
################## Data extract from TMS #######################################

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
browser.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td/fieldset/table/tbody/tr/td[1]/b/img[2]').click() # Click the search

print("test")

while True:
    pass