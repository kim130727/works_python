#KISS에서 url 불러온 후 DBPIA 데이터 수집 농약과학회 등 복잡한 url 구조 (버그 검색)
from selenium import webdriver
import time
from bs4 import BeautifulSoup

# webdirver옵션에서 headless기능을 사용하겠다 라는 내용
webdriver_options = webdriver.ChromeOptions()
webdriver_options .add_argument('headless')

# 페이지불러오기
driver = webdriver.Chrome('C:\projects\\rpa\web crawling\chromedriver.exe', options=webdriver_options)
driver.implicitly_wait(2)

def xpath(site):
    url = "http://kiss.kstudy.com/journal/journal-view.asp?key1=1757&key2=6016"
    driver.get(url)
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="contents"]/div/div/div[1]/ul/li[2]/a').click()
    time.sleep(2)
    print (site)
    driver.find_element_by_xpath(str(site)).click()  # 25권 1호
    time.sleep(5)
    for number1 in range(1, 20):
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        f = open("c:\data automation\\machine2021.txt", 'a', encoding='UTF8')
        number = 1
        while number < 30:
            driver.find_element_by_xpath('//*[@id ="form_main"]/table/tbody/tr['+str(number)+']/td[2]/div/h5/a').click()
            div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div/section[2]/h3')
            print("Title ", div_elems[0].text)
            f.write("Title "+div_elems[0].text)
            f.write(" , ")
            div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div/section[2]/div[2]/div[1]/div')
            print("Author ", div_elems[0].text)
            f.write("Author ", div_elems[0].text)
            f.write(" , ")
            div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div/section[2]/div[2]/div[1]/ul/li[2]')
            print("journal ", div_elems[0].text)
            f.write("journal ", div_elems[0].text)
            f.write(" , ")
            div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div/section[2]/div[2]/div[1]/ul/li[4]')
            print("date ", div_elems[0].text)
            f.write("date ", div_elems[0].text)
            f.write(" , ")
            div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div/div[1]/div[1]/section[3]/div')
            print("keyword ", div_elems[0].text)
            f.write("keyword ", div_elems[0].text)
            f.write(" , ")
            div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div/div[1]/div[1]/section[4]/div[1]')
            print("abstract ", div_elems[0].text)
            f.write("abstract", div_elems[0].text)
            f.write("[Enter")
            #//*[@id="contents"]/div/section[2]/div[2]/div[1]/ul/li[2] #journal
            #//*[@id="contents"]/div/section[2]/div[2]/div[1]/ul/li[4] #date
            #//*[@id="contents"]/div/div[1]/div[1]/section[3]/div   #keyword
            #//*[@id="contents"]/div/div[1]/div[1]/section[4]/div[1] #abstract
            number = number + 3
            # f = open("c:\data automation\\machine2021.txt", 'a')
            # f.write(title[0].text)
            driver.back()
            time.sleep(3)
        f.close()
        number2 = number1 + 1
        try:
            driver.find_element_by_xpath('//*[@id="contents"]/div/div[2]/a[' + str(number2) + ']').click()  # next button
        except:
            break
        # //*[@id="contents"]/div/div[2]/a[2]
        # //*[@id="contents"]/div/div[2]/a[11]
        # //*[@id="contents"]/div/div[2]/a[12]
        time.sleep(5)
    return site


site_22_2 = '//*[@id="addView"]/tr[2]/td[2]/p[1]/a'

print (xpath(site_22_2))