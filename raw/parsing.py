import requests
import re
from bs4 import BeautifulSoup
import codecs

file = codecs.open('c:\data automation\data5_2.txt', 'r', 'utf-8')
lines = file.readlines()

read = re.compile('PD........')
read2 = re.compile('PD\d+\w+\d+')
read3 = re.compile('LS\d+')
read4 = re.compile('WP\d+')
readall= read.findall(str(lines))
read2all = read2.findall(str(lines))
read3all = read3.findall(str(lines))
read4all = read4.findall(str(lines))
totalread = readall + read2all + read3all + read4all

print(totalread)
file.close()

n = 0

while n < 100000:

    try:
        regno = totalread[n]
        n += 1

        url = 'http://www.chinapesticide.gov.cn/myquery/querydetail_en?pdno='+ regno
        source_code = requests.get(url)
        plain_text = source_code.text
        soup = BeautifulSoup(plain_text, 'lxml')

        form = re.compile('<td class="tab_lef_bot_rig" width="30">..')
        comp = re.compile('<td class="tab_lef_bot" width="180">.....')
        per = re.compile('<td class="tab_lef_bot_rig" colspan="3">......................')

        formulation = form.findall(str(soup))
        company = comp.findall(str(soup))
        period = per.findall(str(soup))

        try:
            print(formulation[0][39:])
            print(company[0][36:])
            print(period[0][40:])

            f = codecs.open('c:\data automation\china_registration4.txt', 'a', 'utf-8')
            f.write(regno)
            f.write('  ')
            f.write(formulation[0][39:])
            f.write('  ')
            f.write(period[0][40:])
            f.write('  ')
            f.write('\n')
            f.close()

        except IndexError:
            formulation = ""
            company = ""
            period = ""
            pass

    except IndexError:
        break