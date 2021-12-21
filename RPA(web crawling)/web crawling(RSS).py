#RSS에서 url 불러온 후 엑셀로 보내기까지
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
from bs4 import BeautifulSoup

# 페이지불러오기
driver = webdriver.Chrome('C:\projects\\rpa\web crawling\chromedriver.exe')
driver.implicitly_wait(2)

def url(site):
    driver.get(site)
    time.sleep(5)
    #driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[3]/div[1]/div[2]/div/div[1]/label').click()
    #time.sleep(1)
    #driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[3]/div[1]/div[2]/div/div[2]/div/ul/li[5]/a').click()
    #time.sleep(1)
    #driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[3]/div[1]/div[2]/button').click()
    #time.sleep(3)
    try:
        driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[3]/div[1]/div[1]/label/span').click()
        #'//*[@id="soptionview"]/div/div[4]/div[1]/div[1]/label/span'
        #'//*[@id="soptionview"]/div/div[4]/div[1]/div[1]/ul/li[1]/a'
        print ('click1')
        driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[3]/div[1]/div[1]/ul/li[1]/a').send_keys('\n')
    except:
        driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[4]/div[1]/div[1]/label/span').click()
        print('click2')
        driver.find_element_by_xpath('//*[@id="soptionview"]/div/div[4]/div[1]/div[1]/ul/li[1]/a').send_keys('\n')

    time.sleep(3)
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_xpath('//*[@id="wrap"]/form/div/div[2]/div[1]/div/ul/li[3]/label').click()
    driver.find_element_by_xpath('//*[@id="riss_gubun"]/ul/li[2]/label').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="excel_gubun"]/ul/li[2]/label').click()
    driver.find_element_by_xpath('//*[@id="riss_gubun"]/div[4]/a[1]').send_keys('\n')
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return site

url1 = '  ' #주소 입력
print (url(url1))