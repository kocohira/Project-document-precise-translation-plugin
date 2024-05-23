from selenium import webdriver
from time import sleep

driver = webdriver.Chrome('D:\应用方面\python project2\chromedriver.exe')
driver.get('https://www.baidu.com/')

element = driver.find_element_by_id('kw')

button = element.send_keys('焼肉やる気')

driver.find_element_by_id("su").click()

element.clear()

sleep(10)

driver.quit()
###NT