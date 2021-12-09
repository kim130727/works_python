## parser.py
import requests
from bs4 import BeautifulSoup

from urllib.request import Request, urlopen
from bs4 import BeautifulSoup

base_url = 'https://www.ibuybeauti.com/'
req = Request('https://www.ibuybeauti.com/')
res = urlopen(req)
html = res.read()

bs = BeautifulSoup(html, 'html.parser')
tags = bs.findAll('li', attrs={'class': 'image-wrap'})

for tag in tags :
    # 검색된 태그에서 a 태그에서 텍스트를 가져옴

    print (tag.find('img')['alt'])
    print(tag.find('a')['href'])
    print (tag.find('img')['src'])





# 인덱스를 주고 싶다면 enumerate를 사용한다.
# for index, tag in enumerate(tags):
#    print(str(index) + " : " + tag.a.text)
