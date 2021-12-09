## parser.py
import requests
from bs4 import BeautifulSoup

from urllib.request import Request, urlopen
from bs4 import BeautifulSoup

req = Request('https://www.ibuybeauti.com/')
res = urlopen(req)
html = res.read()

bs = BeautifulSoup(html, 'html.parser')
tags = bs.findAll('li', attrs={'class': 'title'})

for tag in tags :
    # 검색된 태그에서 a 태그에서 텍스트를 가져옴
    print (tag)
    print (tag.find('a').text)
    print (tag.find('a')['href'])

# 인덱스를 주고 싶다면 enumerate를 사용한다.
# for index, tag in enumerate(tags):
#    print(str(index) + " : " + tag.a.text)
