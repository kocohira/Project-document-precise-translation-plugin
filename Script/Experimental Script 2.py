from selenium import webdriver
import time




#实例化谷歌设置选项
option = webdriver.ChromeOptions()
#添加保持登录的数据路径：安装目录一般在C:\Users\黄\AppData\Local\Google\Chrome\User Data
option.add_argument(r'user-data-dir=C:\Users\kocohira\AppData\Local\Google\Chrome\User Data2')

#初始化driver
driver = webdriver.Chrome(options=option)





driver = webdriver.Chrome('D:\应用方面\python project2\chromedriver.exe')
#driver.find_element_by_partial_link_text('美团旅行').click()
driver.get('https://eo.dianping.com/epassport/bookauthmanage')
'''driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'https')]"))


driver.find_element_by_id('login').send_keys('fancy8133')

driver.find_element_by_id('password').send_keys('jack123456')

#driver.switch_to.default_content(driver.find_element_by_class_name('login__submit btn btn_primary btn_m').click())
driver.find_element_by_xpath("//*[@id='login-form']").click()








element = driver.find_element_by_class_name('ant-input')

button = element.send_keys('78735286')

driver.find_element_by_class_name("ant-btn ant-btn-primary").click()

element.clear()

sleep(3)

driver.quit()'''
###NT
