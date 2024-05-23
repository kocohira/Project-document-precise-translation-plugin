#coding=utf-8 
from selenium import webdriver 
from selenium.webdriver.common.keys import Keys #需要引入 keys 包
import os,time

driver = webdriver.Chrome()
driver.get("http://passport.kuaibo.com/login/?referrer=http%3A%2F%2Fwebcloud .kuaibo.com%2F")

time.sleep(3) 
driver.maximize_window() # 浏览器全屏显示

driver.find_element_by_id("user_name").clear() 
driver.find_element_by_id("user_name").send_keys("fnngj")

#tab 的定位相相于清除了密码框的默认提示信息，等同上面的 clear() 
'''driver.find_element_by_id("user_name").send_keys(Keys.TAB) 
time.sleep(3) 
driver.find_element_by_id("user_pwd").send_keys("123456")'''

#通过定位密码框，enter（回车）来代替登陆按钮
driver.find_element_by_id("user_pwd").send_keys(Keys.ENTER)

'''#也可定位登陆按钮，通过 enter（回车）代替 click() 
driver.find_element_by_id("login").send_keys(Keys.ENTER) '''
time.sleep(3)

driver.quit()
###NT


