
from urllib import parse

search = parse.urlparse('http://www.boannews.com/search/news_list.asp?search=title&find=취약점')
print (search)
query = parse.parse_qs(search.query)
print (query)
S_query = parse.urlencode(query, encoding ='euc-kr', doseq = True)
print (S_query)
url = "https://www.boannews.com/search/news_list.asp?{}".format(S_query)
print (url)

import requests
from bs4 import BeautifulSoup
from collections import OrderedDict

news_link =[]
response = requests.get(url)
html = response.text
soup = BeautifulSoup(html, 'html.parser')

bs = BeautifulSoup(html, 'html.parser')
tags = bs.findAll('li', attrs={'class': 'title'})