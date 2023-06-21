##excel에서 pandas - numpy로 데이터 옮긴 후 xml로 변환

import pandas as pd
import re

filename = "C:\Python\RPA\안치현db.xlsx"
df = pd.read_excel(filename)

n = 0
while n < 5:
    print (df.values[n][0])
    print (df.values[n][1])
    print (df.values[n][2])
    print (df.values[n][3])
    print (df.values[n][4])

    text0001 = df.values[n][0]
    text0002 = df.values[n][1]
    text0003 = df.values[n][2]
    text0004 = df.values[n][3]
    text0005 = df.values[n][4]

    filename = "C:\Python\RPA\\230621_sample_0001.xml"
    f = open(filename, 'r', encoding='UTF8')
    lines = f.readlines()
    parse = lines[0]

    n1 = 0
    try:
        while n1 < 500000:
            parse = lines[n1]
            parse = re.sub('TEXT_0001', text0001, parse)
            parse = re.sub('TEXT_0002', text0002, parse)
            parse = re.sub('TEXT_0003', text0003, parse)
            parse = re.sub('TEXT_0004', text0004, parse)
            parse = re.sub('TEXT_0005', text0005, parse)
            f2 = open('C:\Python\RPA\\final\\final'+str(n)+'.xml', 'a', encoding='UTF8')
            f2.write(parse)
            f2.close()
            n1 = n1+1
    except:
        pass
    n = n+1