## parser.py
import requests
from bs4 import BeautifulSoup
import json
import os

## python파일의 위치
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

req = requests.get('http://farmhannong.esafe.or.kr/front/Contents.do?cmd=contentsFrame&p_userid=farm_1089993&p_contsid=0002019&p_isclosed=N&p_eduprocess=S&p_subj=4328&p_year=2019&p_subjseq=0015&p_from=userlist&p_itemid=1&p_directitemid=1&p_isback=&p_isgoyong=N', verify=False)
html = req.text
soup = BeautifulSoup(html, 'html.parser')
my_titles = soup.select(
    'h3 > a'
    )

data = {}

for title in my_titles:
    data[title.text] = title.get('href')

with open(os.path.join(BASE_DIR, 'result.json'), 'w+') as json_file:
    json.dump(data, json_file)