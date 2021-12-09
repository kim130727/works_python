#excel에서 pandas - numpy로 데이터 옮긴 후 xml로 변환

import pandas as pd
import re

filename = "../Sample_excel.xls"
df = pd.read_excel(filename)

print (df)

data = df[df['시험번호'] == '20-RC001']
print ("data", data)
print (data.values[0][0])
print (data.values[0][1])
print (data.values[0][2])
print (data.values[0][3])
print (data.values[0][4])
print (data.values[0][5].strftime('%Y'+"-"+"%m"+"-"+"%d"))
print (data.values[0][6])
print (data.values[0][7])
print (data.values[0][8])
print (data.values[0][9])
print (data.values[0][10])
print (data.values[0][11])
print (data.values[0][12])
print (data.values[0][13])
print (data.values[0][14])
print (data.values[0][15])
print (data.values[0][16])
print (data.values[0][17])
print (data.values[0][18])
print (data.values[0][19])

시험번호 = data.values[0][0]
시험책임자 = data.values[0][1]
시험물질 = data.values[0][2]
분석성분 = data.values[0][3]
작물 = (data.values[0][4])
계획서승인일 = str(data.values[0][5].strftime('%Y'+"-"+"%m"+"-"+"%d"))
적용병해충 = (data.values[0][6])
병해충발생시기 = (data.values[0][7])
사용적기및방법 = (data.values[0][8])

filename = "../21년시험계획서.xml"

f = open(filename, 'r', encoding='UTF8')
lines = f.readlines()

print(lines)

n = 0
try:
    while n < 500000:
        parse = lines[n]
        parse = re.sub('{{시험번호}}', 시험번호, parse)
        parse = re.sub('{{시험책임자}}', 시험책임자, parse)
        parse = re.sub('{{시험물질}}', 시험물질, parse)
        parse = re.sub('{{분석성분}}', 분석성분, parse)
        parse = re.sub('{{작물}}', 작물, parse)
        parse = re.sub('{{계획서승인일}}', 계획서승인일, parse)
        parse = re.sub('{{적용병해충}}', 적용병해충, parse)
        parse = re.sub('{{병해충발생시기}}', 병해충발생시기, parse)
        parse = re.sub('{{사용적기및방법}}', 사용적기및방법, parse)
        print ("checking")
        print (parse)
        f2 = open('../final.xml', 'a', encoding='UTF8')
        f2.write(parse)
        f2.close()
        n = n+1
except:
    pass
f.close()