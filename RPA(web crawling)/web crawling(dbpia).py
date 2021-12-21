#DBPIA 데이터 수집 농약과학회 등 복잡한 url 구조 (버그 검색)
#headliss 적용
#초록 및 키워드 포함

from selenium import webdriver
import time

# webdirver옵션에서 headless기능을 사용하겠다 라는 내용
webdriver_options = webdriver.ChromeOptions()
webdriver_options .add_argument('headless')

# 페이지불러오기
driver = webdriver.Chrome('C:\projects\\rpa\web crawling\chromedriver.exe', options=webdriver_options)
driver.implicitly_wait(2)

def xpath(site):
    print (site)
    driver.get(site)
    time.sleep(1)
    number2 = 1
    while number2 < 31:
        try:
            for x in range(1,10):
                driver.find_element_by_xpath('//*[@id="contents"]/div[2]/div[2]/div[2]/div[2]/div/a/span').click()
                time.sleep(0.2)
        except:
            print ("try1")
            pass
        try:
            driver.find_element_by_xpath('//*[@id="voisNodeList"]/ul[2]/li[' + str(number2) + ']/div/div[2]/h5/a').click()
            # //*[@id="voisNodeList"]/ul[2]/li[2]/div/div[2]/h5/a
            # //*[@id="voisNodeList"]/ul/li[2]/div/div[2]/h5/a
            print ("click")
            time.sleep(1)
        except:
            try:
                driver.find_element_by_xpath(
                    '//*[@id="voisNodeList"]/ul/li[' + str(number2) + ']/div/div[2]/h5/a').click()
                # //*[@id="voisNodeList"]/ul[2]/li[2]/div/div[2]/h5/a
                # //*[@id="voisNodeList"]/ul/li[2]/div/div[2]/h5/a
                print("click")
                time.sleep(1)
            except:
                print("search other list")
                pass
            pass

        try:
            driver.find_element_by_xpath('//*[@id="#pub_modalOrganPop"]').click()
        except:
            pass

        f = open("c:\data automation\\data_raw.txt", 'a', encoding='UTF8')
        div_elems = driver.find_elements_by_xpath('//*[@id="dev_node_title"]')  # Title
        # //*[@id="dev_node_title"]
        # //*[@id="contents"]/div[2]/div[1]/div
        try:
            print("1 ", div_elems[0].text)
        except:
            print ("next url")
            break
        f.write(div_elems[0].text)
        f.write(' , ')
        div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div[2]/div[1]/div/div[2]/p')
        print("2 ", div_elems[0].text)
        f.write(div_elems[0].text)
        f.write(' , ')
        div_elems = driver.find_elements_by_xpath('//*[@id="contents"]/div[2]/div[1]/div/div[2]/ul')
        print("3 ", div_elems[0].text)
        f.write(div_elems[0].text)
        f.write(' , ')
        try:
            div_elems = driver.find_elements_by_xpath('//*[@id="pub_abstract"]/div[2]/div/p[1]')
            print("4 ", str(div_elems[0].text))
            f.write(' , ')
        except:
            pass
        try:
            div_elems = driver.find_elements_by_xpath('//*[@id="pub_abstract"]/div[2]/div/div[1]/div[1]')
            print("4 ", str(div_elems[0].text))
            f.write(div_elems[0].text)
        except:
            pass

        div_elems = driver.find_elements_by_xpath('// *[ @ id = "pub_keyword"] / div / dl')
        try:
            print("5 ", div_elems[0].text)
            f.write(div_elems[0].text)
        except:
            pass
        f.write('[enter]  ')
        time.sleep(3)
        f.close()
        driver.get(site)
        # driver.back()
        time.sleep(3)
        number2 = number2 + 1
        print('number2: ', number2)
    return site

url1 = 'https://www.dbpia.co.kr/journal/voisDetail?voisId=VOIS00635903'

print (xpath(url1))