from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as pd
import time

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome('./chromedriver.exe',options=options)
data = driver.get('http://www.nepalstock.com/floorsheet')
sel = Select(driver.find_element_by_id('stock-symbol'))
time.sleep(2)
sel.select_by_visible_text('NBB') #Change Scrip Name here 
time.sleep(2)
selrow = Select(driver.find_element_by_xpath('//*[@id="news_info-filter"]/label[5]/select'))
time.sleep(2)
selrow.select_by_visible_text('500')
time.sleep(2)
driver.find_element_by_xpath('//*[@id="news_info-filter"]/input[1]').click()
time.sleep(2)
total = []
for i in range(3,503):
    data = {
        'buyer': driver.find_element_by_xpath(f'/html/body/div[5]/table/tbody/tr[{i}]/td[4]').text,
        'seller': driver.find_element_by_xpath(f'//*[@id="home-contents"]/table/tbody/tr[{i}]/td[5]').text,
        'quantity': driver.find_element_by_xpath(f'//*[@id="home-contents"]/table/tbody/tr[{i}]/td[6]').text,
        'price': driver.find_element_by_xpath(f'//*[@id="home-contents"]/table/tbody/tr[{i}]/td[7]').text,
        'amount': driver.find_element_by_xpath(f'//*[@id="home-contents"]/table/tbody/tr[{i}]/td[8]').text,
    }
    total.append(data)

# print(total)

df = pd.DataFrame(total)
df.to_excel('./floorsheet.xlsx',index=False)
driver.close()

