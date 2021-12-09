# -*- coding: utf-8 -*-
import pypyodbc
import os
import re
import win32com.client
import win32timezone
import pubchempy as pcp
import codecs
import numpy as np
import sys
import glob
import os.path

nm = input('시험번호를 입력해 주세요 ->')
mn = input('번호를 입력해 주세요. 1. 이화학, 2. P1, 3. P2  ->')

if mn == '1':
    ope = '이화학'
elif mn == '2':
    ope = 'P1'
elif mn == '3':
    ope = 'P2'
else:
    pass

pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\data automation\이화학적분석.accdb;")

c = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c.execute(sql, (nm, ope))
r = c.fetchone()

print(r)
print(r[0])

구분 = r[2]
시험내용 = r[10]
시료명1 = r[18]
한글명1 = r[19]
시료명2 = r[38]
한글명2 = r[39]
시료명3 = r[58]
한글명3 = r[59]
시험물질 = r[3]
분석일 = r[9]
print (분석일)
print (type(분석일))
print ('테스트', re.findall(r"(\d+)", 분석일[0]))
분석년월일 = list(re.findall(r"(\d+)", 분석일))[0] + "년 " + list(re.findall(r"(\d+)", 분석일))[1] + "월 " + list(re.findall(r"(\d+)", 분석일))[2] + "일"
시료1_1 = r[23]
시료1_2 = r[27]
시료1_3 = r[31]
시료1평균 = r[32]
시료2_1 = r[43]
시료2_2 = r[47]
시료2_3 = r[51]
시료2평균 = r[52]
시료3_1 = r[63]
시료3_2 = r[67]
시료3_3 = r[71]
시료3평균 = r[72]
분석기기1 = r[11]
분석기기2 = r[12]
분석기기3 = r[13]
성상 = r[74]
색상 = r[75]
냄새 = r[76]
시험내용 = r[10]

if 시험내용 == "P1":
    시험항목 = "이화학시험, 약효/약해시험, 지역적응성시험"
elif 시험내용 == "P2":
    시험항목 = "이화학시험, 약효/약해시험, 잔류시험, 독성시험, 지역적응성시험"
elif 시험내용 == "이화학":
    시험항목 = "이화학시험"
else:
    시험항목 = ""

if 성상 == None:
    성상 = ""
else:
    pass

if 색상 == None:
    색상 = ""
else:
    pass

if 냄새 == None:
    냄새 = ""
else:
    pass

if 시료명2 == None:
    한글명2 = ""
    시료2_1 = ""
    시료2_2 = ""
    시료2_3 = ""
    시료2 = ""
    시료2평균 = ""
    시료2std = ""
else:
    시료2 = np.array([float(시료2_1), float(시료2_2), float(시료2_3)])
    시료2std = round(np.std(시료2, axis=0, ddof=1), 5)

if 시료명3 == None:
    한글명3 = ""
    시료3_1 = ""
    시료3_2 = ""
    시료3_3 = ""
    시료3 = ""
    시료3평균 = ""
    시료3std = ""
else:
    시료3 = np.array([float(시료3_1), float(시료3_2), float(시료3_3)])
    시료3std = round(np.std(시료3, axis=0, ddof=1), 5)

시료1 = np.array([float(시료1_1), float(시료1_2), float(시료1_3)])
시료1std = round(np.std(시료1, axis=0, ddof=1), 5)

책임자 = r[7]
의뢰자 = r[6]

if 분석기기1 == 분석기기2:
    분석기기2 = ""

elif 분석기기2 == None:
    분석기기2 = ""

else:
    pass

if 분석기기1 == 분석기기3:
    분석기기3 = ""

elif 분석기기3 == None:
    분석기기3 = ""

else:
    pass

try:
    if 시료명1 == None:
        시료명1IUPAC = ""

    else:
        시료명1CID = pcp.get_compounds(시료명1, 'name')
        시료명1IUPAC = 시료명1CID[0].iupac_name

except IndexError:
    시료명1IUPAC = ""

try:
    if 시료명2 == None:
        시료명2IUPAC = ""
        시료명2 = ""

    else:
        시료명2CID = pcp.get_compounds(시료명2, 'name')
        시료명2IUPAC = 시료명2CID[0].iupac_name

except IndexError:
    시료명2IUPAC = ""

try:
    if 시료명3 == None:
        시료명3IUPAC = ""
        시료명3 = ""

    else:
        시료명3CID = pcp.get_compounds(시료명3, 'name')
        시료명3IUPAC = 시료명3CID[0].iupac_name

except IndexError:
    시료명3IUPAC = ""

제형코드 = 시험물질[-2:]

LOTNO = r[4]

try:
    제조년 = '20' + list(re.findall(r"(\d+)", LOTNO)[0])[0] + list(re.findall(r"(\d+)", LOTNO)[0])[1]
    제조월 = list(re.findall(r"(\d+)", LOTNO)[0])[2] + list(re.findall(r"(\d+)", LOTNO)[0])[3]
    제조일 = list(re.findall(r"(\d+)", LOTNO)[0])[4] + list(re.findall(r"(\d+)", LOTNO)[0])[5]
except IndexError:
    제조년 = ""
    제조월 = ""
    제조일 = ""

# 제형코드는 DT 하나만 설정했음 나머지도 해야함.

검사항목1 = ""
검사항목2 = ""
검사항목3 = ""
검사항목4 = ""
검사항목5 = ""

if 제형코드 == "EC":
    제형분류 = "유제"
    수량 = "500ml"
    검사항목1 = "유화성: " + r[77]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "SL":
    제형분류 = "액제"
    검사항목1 = "수용성: " + r[80]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    수량 = "500ml"

elif 제형코드 == "SC":
    제형분류 = "액상수화제"
    수량 = "500ml"
    검사항목1 = "수화성: " + r[78]
    검사항목2 = "분말도: " + r[81]
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "WP":
    제형분류 = "수화제"
    수량 = "500g"
elif 제형코드 == "SP":
    제형분류 = "수용제"
    수량 = "500g"
elif 제형코드 == "OD":
    제형분류 = "유상수화제"
    수량 = "500ml"
elif 제형코드 == "EP":
    제형분류 = "분상유제"
    수량 = "500g"
elif 제형코드 == "DP":
    제형분류 = "분제"
    수량 = "500g"
    검사항목1 = "분말도: " + r[81]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "DS":
    제형분류 = "분의제"
    수량 = "500g"
elif 제형코드 == "GP":
    제형분류 = "미분제"
    수량 = "500g"
elif 제형코드 == "DL":
    제형분류 = "저비산분제"
    수량 = "500g"
elif 제형코드 == "GR":
    제형분류 = "입제"
    수량 = "500g"
elif 제형코드 == "MG":
    제형분류 = "미립제"
    수량 = "500g"
elif 제형코드 == "PA":
    제형분류 = "도포제"
    수량 = "500ml"
elif 제형코드 == "GA":
    제형분류 = "훈증제"
    수량 = "500ml"
elif 제형코드 == "FU":
    제형분류 = "훈연제"
    수량 = "500ml"
elif 제형코드 == "AE":
    제형분류 = "연무제"
    수량 = "500ml"
elif 제형코드 == "CG":
    제형분류 = "캡슐제"
    수량 = "500ml"
elif 제형코드 == "FG":
    제형분류 = "세립제"
    수량 = "500g"
elif 제형코드 == "FG":
    제형분류 = "세립제"
    수량 = "500g"
elif 제형코드 == "FW":
    제형분류 = "과립훈연제"
    수량 = "500g"
    검사항목1 = "발연성: " + r[94]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "WF":
    제형분류 = "수화성미분제"
    수량 = "500g"
elif 제형코드 == "WG":
    제형분류 = "입상수화제"
    수량 = "500g"
    검사항목1 = "수화성: " + r[78]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "EW":
    제형분류 = "유탁제"
    수량 = "500ml"
    검사항목1 ="유화성: " + r[77]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "CS":
    제형분류 = "캡슐현탁제"
    수량 = "500ml"
elif 제형코드 == "SE":
    제형분류 = "유현탁제"
    수량 = "500ml"
    검사항목1 = "유화성: " + r[77]
    검사항목2 = "수화성: " + r[78]
    검사항목3 = "분말도: " + r[81]
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "DC":
    제형분류 = "분산성액제"
    수량 = "500ml"
    검사항목1 = "수중분산성: " + r[79]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "SO":
    제형분류 = "수면전개제"
    수량 = "500ml"
elif 제형코드 == "WS":
    제형분류 = "종자처리수화제"
    수량 = "500g"
elif 제형코드 == "ME":
    제형분류 = "미탁제"
    수량 = "500ml"
elif 제형코드 == "FS":
    제형분류 = "종자처리액상수화제"
    수량 = "500ml"
elif 제형코드 == "UG":
    제형분류 = "수면부상성입제"
    수량 = "500g"
elif 제형코드 == "PF":
    제형분류 = "비닐멀칭제"
    수량 = "500g"
elif 제형코드 == "SF":
    제형분류 = "판상줄제"
    수량 = "500g"
elif 제형코드 == "OL":
    제형분류 = "오일제"
    수량 = "500ml"
elif 제형코드 == "SG":
    제형분류 = "입상수용제"
    수량 = "500g"
elif 제형코드 == "AL":
    제형분류 = "직접살포액제"
    수량 = "500ml"
elif 제형코드 == "DT":
    제형분류 = "직접살포정제"
    수량 = "500g"
elif 제형코드 == "VP":
    제형분류 = "마이크로캡슐훈증제"
    수량 = ""
elif 제형코드 == "AS":
    제형분류 = "액상제"
    수량 = "500ml"
elif 제형코드 == "EM":
    제형분류 = "유상현탁제"
    수량 = "500ml"
elif 제형코드 == "SM":
    제형분류 = "액상현탁제"
    수량 = "500ml"
elif 제형코드 == "GM":
    제형분류 = "고상제"
    수량 = "500g"
elif 제형코드 == "GG":
    제형분류 = "대립제"
    수량 = "500g"
    검사항목1 = "확산성: " + r[93]
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
elif 제형코드 == "WT":
    제형분류 = "정제상수화제"
    수량 = "500g"
elif 제형코드 == "ZC":
    제형분류 = "캡슐액상수화제"
    수량 = "500ml"
elif 제형코드 == "GD":
    제형분류 = "발생기"
    수량 = ""
elif 제형코드 == "RM":
    제형분류 = "발생제"
    수량 = ""
else:
    제형분류 = ""
    수량 = ""

print(시험물질)

시험물질_pre = re.sub(r"([ (+)%])+", "             ", 시험물질)

print('시험물질pre', 시험물질_pre)
함량 = re.findall(r"(\d\d?.?\d?\d?\d?)", 시험물질_pre)

함량.extend(["", "", ""])
print('함량', 함량)

소수점1 = re.findall(r"([.]\d+)", 함량[0])
소수점2 = re.findall(r"([.]\d+)", 함량[1])
소수점3 = re.findall(r"([.]\d+)", 함량[2])

if len(str(소수점1)) == 2:
    point1 = 2
else:
    point1 = len(str(소수점1)) - 3

if len(str(소수점2)) == 2:
    point2 = 2
else:
    point2 = len(str(소수점2)) - 3

if len(str(소수점3)) == 2:
    point3 = 2
else:
    point3 = len(str(소수점3)) - 3

print('소수점1', 소수점1, len(str(소수점1)))
print('소수점2', 소수점2, len(str(소수점2)))
print('소수점3', 소수점3, len(str(소수점3)))
print('point1', point1)
print('point2', point2)
print('point3', point3)

# print (함량)으로 작성하면 결과가 ['8', '3', '1', '', ''] 이렇게 나옴? 이유를 모르겠음.

함량1 = 함량[0] + '%'

if 함량[1] == "":
    함량2 = ""
else:
    함량2 = 함량[1] + '%'

if 함량[2] == "":
    함량3 = ""
else:
    함량3 = 함량[2] + '%'

std_content1 = r[14]
std_g1 = r[15]
std_AI_area1 = r[16]
std_IS_area1 = r[17]

sam_1_1_g = r[20]
sam_1_1_AI = r[21]
sam_1_1_IS = r[22]

sam_1_2_g = r[24]
sam_1_2_AI = r[25]
sam_1_2_IS = r[26]

sam_1_3_g = r[28]
sam_1_3_AI = r[29]
sam_1_3_IS = r[30]

factor_std1 = float(std_g1) * float(std_content1) * float(std_IS_area1) / float(std_AI_area1)
sam_1_1_content = round((factor_std1 * float(sam_1_1_AI)) / (float(sam_1_1_IS) * (float(sam_1_1_g))), int(point1))
sam_1_2_content = round((factor_std1 * float(sam_1_2_AI)) / (float(sam_1_2_IS) * (float(sam_1_2_g))), int(point1))
sam_1_3_content = round((factor_std1 * float(sam_1_3_AI)) / (float(sam_1_3_IS) * (float(sam_1_3_g))), int(point1))
sam_1_average = round((sam_1_1_content + sam_1_2_content + sam_1_3_content) / 3, int(point1))
sam_1_stdev = round(((((sam_1_1_content - ((sam_1_1_content + sam_1_2_content + sam_1_3_content) / 3)) ** 2 + (
sam_1_2_content - ((sam_1_1_content + sam_1_2_content + sam_1_3_content) / 3)) ** 2 + (sam_1_3_content - (
(sam_1_1_content + sam_1_2_content + sam_1_3_content) / 3)) ** 2)) / 2) ** 0.5, 5)

std_content2 = r[34]
std_g2 = r[35]
std_AI_area2 = r[36]
std_IS_area2 = r[37]

sam_2_1_g = r[40]
sam_2_1_AI = r[41]
sam_2_1_IS = r[42]

sam_2_2_g = r[44]
sam_2_2_AI = r[45]
sam_2_2_IS = r[46]

sam_2_3_g = r[48]
sam_2_3_AI = r[49]
sam_2_3_IS = r[50]

if 시료명2 == "":
    factor_std2 = ""
    sam_2_1_content = ""
    sam_2_2_content = ""
    sam_2_3_content = ""
    sam_2_average = ""
    sam_2_stdev = ""

else:
    factor_std2 = float(std_g2) * float(std_content2) * float(std_IS_area2) / float(std_AI_area2)
    sam_2_1_content = round((factor_std2 * float(sam_2_1_AI)) / (float(sam_2_1_IS) * (float(sam_2_1_g))), int(point2))
    sam_2_2_content = round((factor_std2 * float(sam_2_2_AI)) / (float(sam_2_2_IS) * (float(sam_2_2_g))), int(point2))
    sam_2_3_content = round((factor_std2 * float(sam_2_3_AI)) / (float(sam_2_3_IS) * (float(sam_2_3_g))), int(point2))
    sam_2_average = round((sam_2_1_content + sam_2_2_content + sam_2_3_content) / 3, int(point2))
    sam_2_stdev = round(((((sam_2_1_content - ((sam_2_1_content + sam_2_2_content + sam_2_3_content) / 3)) ** 2 + (
    sam_2_2_content - ((sam_2_1_content + sam_2_2_content + sam_2_3_content) / 3)) ** 2 + (sam_2_3_content - (
    (sam_2_1_content + sam_2_2_content + sam_2_3_content) / 3)) ** 2)) / 2) ** 0.5, 5)

std_content3 = r[54]
std_g3 = r[55]
std_AI_area3 = r[56]
std_IS_area3 = r[57]

sam_3_1_g = r[60]
sam_3_1_AI = r[61]
sam_3_1_IS = r[62]

sam_3_2_g = r[64]
sam_3_2_AI = r[65]
sam_3_2_IS = r[66]

sam_3_3_g = r[68]
sam_3_3_AI = r[69]
sam_3_3_IS = r[70]

if 시료명3 == "":
    factor_std3 = ""
    sam_3_1_content = ""
    sam_3_2_content = ""
    sam_3_3_content = ""
    sam_3_average = ""
    sam_3_stdev = ""

else:
    factor_std3 = float(std_g3) * float(std_content3) * float(std_IS_area3) / float(std_AI_area3)
    sam_3_1_content = round((factor_std3 * float(sam_3_1_AI)) / (float(sam_3_1_IS) * (float(sam_3_1_g))), int(point3))
    sam_3_2_content = round((factor_std3 * float(sam_3_2_AI)) / (float(sam_3_2_IS) * (float(sam_3_2_g))), int(point3))
    sam_3_3_content = round((factor_std3 * float(sam_3_3_AI)) / (float(sam_3_3_IS) * (float(sam_3_3_g))), int(point3))
    sam_3_average = round((sam_3_1_content + sam_3_2_content + sam_3_3_content) / 3, int(point3))
    sam_3_stdev = round(((((sam_3_1_content - ((sam_3_1_content + sam_3_2_content + sam_3_3_content) / 3)) ** 2 + (
    sam_3_2_content - ((sam_3_1_content + sam_3_2_content + sam_3_3_content) / 3)) ** 2 + (sam_3_3_content - (
    (sam_3_1_content + sam_3_2_content + sam_3_3_content) / 3)) ** 2)) / 2) ** 0.5, 5)

print('구분: ', 구분)
print('분석년월일: ', 분석년월일)
print('시험내용: ', 시험내용)
print('시험물질: ', 시험물질)
print('시험책임자', 책임자, '시험의뢰자', 의뢰자)
print('시료명: ', 시료명1, 시료명2, 시료명3, 제형코드)
print('IUPAC NAME: ', 시료명1IUPAC)
print('IUPAC NAME: ', 시료명2IUPAC)
print('IUPAC NAME: ', 시료명3IUPAC)
print('한글명: ', 한글명1, 한글명2, 한글명3, 제형분류)
print('수량: ', 수량)
print('함량: ', 함량1, 함량2, 함량3)
print('Lot No.: ', LOTNO, ' 제조년월일:', 제조년 + '년', 제조월 + '월', 제조일 + '일')
print('분석일: ', 분석일)
print('시료1: ', 시료1_1, 시료1_2, 시료1_3, 시료1평균)
print('시료2: ', 시료2_1, 시료2_2, 시료2_3, 시료2평균)
print('시료3: ', 시료3_1, 시료3_2, 시료3_3, 시료3평균)
print('분석기기: ', 분석기기1, 분석기기2, 분석기기3)
print(성상, " ", 색상, " ", 냄새)
print('시료1 factor', factor_std1, '시료2 factor', factor_std2, '시료3 factor', factor_std3)
print('시료1 함량:', sam_1_1_content, sam_1_2_content, sam_1_3_content, '평균', sam_1_average, '표준편차', sam_1_stdev)
print('시료2 함량:', sam_2_1_content, sam_2_2_content, sam_2_3_content, '평균', sam_2_average, '표준편차', sam_2_stdev)
print('시료3 함량:', sam_3_1_content, sam_3_2_content, sam_3_3_content, '평균', sam_3_average, '표준편차', sam_3_stdev)
print('해당 폴더에 P1 파일(c:\works\이화학.xml 을 작성중입니다.')

f = codecs.open("C:\data automation\이화학.xml", 'w', 'utf-8')

f.write('<?xml version="1.0"?>\n')
f.write('<?mso-application progid="Excel.Sheet"?>\n')
f.write('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
f.write(' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
f.write(' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
f.write(' xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"\n')
f.write(' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n')
f.write(' xmlns:html="http://www.w3.org/TR/REC-html40">\n')
f.write(' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n')
f.write('  <Author>농업기술연구소</Author>\n')
f.write('  <LastAuthor>LG PC</LastAuthor>\n')
f.write('  <LastPrinted>2015-03-09T04:29:26Z</LastPrinted>\n')
f.write('  <Created>1999-12-11T04:38:33Z</Created>\n')
f.write('  <LastSaved>2016-02-04T13:33:04Z</LastSaved>\n')
f.write('  <Version>12.00</Version>\n')
f.write(' </DocumentProperties>\n')
f.write(' <CustomDocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n')
f.write('  <IVID2F1E1603 dt:dt="string"></IVID2F1E1603>\n')
f.write('  <IVIDC dt:dt="string"></IVIDC>\n')
f.write('  <IVID362F13E8 dt:dt="string"></IVID362F13E8>\n')
f.write('  <IVID3A3618F1 dt:dt="string"></IVID3A3618F1>\n')
f.write('  <IVID15E41318 dt:dt="string"></IVID15E41318>\n')
f.write('  <IVID181914D9 dt:dt="string"></IVID181914D9>\n')
f.write('  <IVID155815FB dt:dt="string"></IVID155815FB>\n')
f.write('  <IVIDD091BF0 dt:dt="string"></IVIDD091BF0>\n')
f.write('  <IVID344CCFFC dt:dt="string"></IVID344CCFFC>\n')
f.write('  <IVID1A7D12ED dt:dt="string"></IVID1A7D12ED>\n')
f.write('  <IVID1B2115FE dt:dt="string"></IVID1B2115FE>\n')
f.write('  <IVID35431BD0 dt:dt="string"></IVID35431BD0>\n')
f.write('  <IVID4637A884 dt:dt="string"></IVID4637A884>\n')
f.write('  <IVID127C14F5 dt:dt="string"></IVID127C14F5>\n')
f.write('  <IVID1834F0DD dt:dt="string"></IVID1834F0DD>\n')
f.write('  <IVID312119E0 dt:dt="string"></IVID312119E0>\n')
f.write('  <IVID1C5812DA dt:dt="string"></IVID1C5812DA>\n')
f.write('  <IVID173907ED dt:dt="string"></IVID173907ED>\n')
f.write('  <IVID1D3F17E2 dt:dt="string"></IVID1D3F17E2>\n')
f.write('  <IVID13451200 dt:dt="string"></IVID13451200>\n')
f.write('  <IVID475611CF dt:dt="string"></IVID475611CF>\n')
f.write('  <IVID302D13DA dt:dt="string"></IVID302D13DA>\n')
f.write('  <IVIDD5915D9 dt:dt="string"></IVIDD5915D9>\n')
f.write('  <IVID17F6384A dt:dt="string"></IVID17F6384A>\n')
f.write('  <IVID3B5A10EA dt:dt="string"></IVID3B5A10EA>\n')
f.write('  <IVID3D0F16E3 dt:dt="string"></IVID3D0F16E3>\n')
f.write('  <IVID30260FFC dt:dt="string"></IVID30260FFC>\n')
f.write('  <IVID2F301BED dt:dt="string"></IVID2F301BED>\n')
f.write('  <IVID2F1117F5 dt:dt="string"></IVID2F1117F5>\n')
f.write('  <IVID121617DE dt:dt="string"></IVID121617DE>\n')
f.write('  <IVID13691AF2 dt:dt="string"></IVID13691AF2>\n')
f.write('  <IVID1A3B0AF0 dt:dt="string"></IVID1A3B0AF0>\n')
f.write('  <IVID373F12DB dt:dt="string"></IVID373F12DB>\n')
f.write('  <IVID274B1CF5 dt:dt="string"></IVID274B1CF5>\n')
f.write('  <IVID2B4E17FA dt:dt="string"></IVID2B4E17FA>\n')
f.write('  <IVID253D11EF dt:dt="string"></IVID253D11EF>\n')
f.write('  <IVID102124BA dt:dt="string"></IVID102124BA>\n')
f.write('  <IVID3D1509D0 dt:dt="string"></IVID3D1509D0>\n')
f.write('  <IVID35641901 dt:dt="string"></IVID35641901>\n')
f.write('  <IVID45E1ED9 dt:dt="string"></IVID45E1ED9>\n')
f.write('  <IVID324113D1 dt:dt="string"></IVID324113D1>\n')
f.write('  <IVID1A2D1903 dt:dt="string"></IVID1A2D1903>\n')
f.write('  <IVID222F6E42 dt:dt="string"></IVID222F6E42>\n')
f.write('  <IVID137012E9 dt:dt="string"></IVID137012E9>\n')
f.write('  <IVID3D4D17F3 dt:dt="string"></IVID3D4D17F3>\n')
f.write('  <IVID2F2214CF dt:dt="string"></IVID2F2214CF>\n')
f.write('  <IVID212812E2 dt:dt="string"></IVID212812E2>\n')
f.write('  <IVID174513DF dt:dt="string"></IVID174513DF>\n')
f.write('  <IVID14481408 dt:dt="string"></IVID14481408>\n')
f.write('  <IVID2E670A05 dt:dt="string"></IVID2E670A05>\n')
f.write('  <IVID2A161305 dt:dt="string"></IVID2A161305>\n')
f.write('  <IVID173E1206 dt:dt="string"></IVID173E1206>\n')
f.write('  <IVID232310EC dt:dt="string"></IVID232310EC>\n')
f.write('  <IVID133D1AE5 dt:dt="string"></IVID133D1AE5>\n')
f.write('  <IVIDF6113D9 dt:dt="string"></IVIDF6113D9>\n')
f.write('  <IVID362E14DB dt:dt="string"></IVID362E14DB>\n')
f.write('  <IVID1F6511DB dt:dt="string"></IVID1F6511DB>\n')
f.write('  <IVID3F1D10E8 dt:dt="string"></IVID3F1D10E8>\n')
f.write('  <IVID144313EE dt:dt="string"></IVID144313EE>\n')
f.write('  <IVID272C0FEF dt:dt="string"></IVID272C0FEF>\n')
f.write('  <IVID240A1504 dt:dt="string"></IVID240A1504>\n')
f.write('  <IVID2E511106 dt:dt="string"></IVID2E511106>\n')
f.write('  <IVID2A6D14EB dt:dt="string"></IVID2A6D14EB>\n')
f.write('  <IVID386F14FA dt:dt="string"></IVID386F14FA>\n')
f.write('  <IVIDA1B07F3 dt:dt="string"></IVIDA1B07F3>\n')
f.write('  <IVID2A6715D8 dt:dt="string"></IVID2A6715D8>\n')
f.write('  <IVID222D19FF dt:dt="string"></IVID222D19FF>\n')
f.write('  <IVID2D4D15EB dt:dt="string"></IVID2D4D15EB>\n')
f.write('  <IVID1A3517F4 dt:dt="string"></IVID1A3517F4>\n')
f.write('  <IVID2B0E1302 dt:dt="string"></IVID2B0E1302>\n')
f.write('  <IVID332E19D7 dt:dt="string"></IVID332E19D7>\n')
f.write('  <IVID22261800 dt:dt="string"></IVID22261800>\n')
f.write('  <IVID325116DE dt:dt="string"></IVID325116DE>\n')
f.write('  <IVID81113D2 dt:dt="string"></IVID81113D2>\n')
f.write('  <IVID1D231201 dt:dt="string"></IVID1D231201>\n')
f.write('  <IVID366A14F0 dt:dt="string"></IVID366A14F0>\n')
f.write('  <IVID316311F9 dt:dt="string"></IVID316311F9>\n')
f.write('  <IVIDE0715F1 dt:dt="string"></IVIDE0715F1>\n')
f.write('  <IVID3B5816EC dt:dt="string"></IVID3B5816EC>\n')
f.write('  <IVID351414F8 dt:dt="string"></IVID351414F8>\n')
f.write('  <IVID2F251AE7 dt:dt="string"></IVID2F251AE7>\n')
f.write('  <IVID2A5E1D03 dt:dt="string"></IVID2A5E1D03>\n')
f.write('  <IVID306310DF dt:dt="string"></IVID306310DF>\n')
f.write('  <IVID266F16CF dt:dt="string"></IVID266F16CF>\n')
f.write('  <IVID307414D1 dt:dt="string"></IVID307414D1>\n')
f.write('  <IVID344B1400 dt:dt="string"></IVID344B1400>\n')
f.write('  <IVID135B1DF5 dt:dt="string"></IVID135B1DF5>\n')
f.write('  <IVID1A3716D3 dt:dt="string"></IVID1A3716D3>\n')
f.write('  <IVIDD1916DB dt:dt="string"></IVIDD1916DB>\n')
f.write('  <IVID11431AF1 dt:dt="string"></IVID11431AF1>\n')
f.write('  <IVID1B2C19F3 dt:dt="string"></IVID1B2C19F3>\n')
f.write('  <IVIDD5E0FE6 dt:dt="string"></IVIDD5E0FE6>\n')
f.write('  <IVID162D1605 dt:dt="string"></IVID162D1605>\n')
f.write('  <IVID28741007 dt:dt="string"></IVID28741007>\n')
f.write('  <IVID2A3614FA dt:dt="string"></IVID2A3614FA>\n')
f.write('  <IVID15231CDF dt:dt="string"></IVID15231CDF>\n')
f.write('  <IVID322814F3 dt:dt="string"></IVID322814F3>\n')
f.write('  <IVID2F6C14EF dt:dt="string"></IVID2F6C14EF>\n')
f.write('  <IVID252617FB dt:dt="string"></IVID252617FB>\n')
f.write('  <IVIDA0D1BD8 dt:dt="string"></IVIDA0D1BD8>\n')
f.write('  <IVID3E4418F8 dt:dt="string"></IVID3E4418F8>\n')
f.write('  <IVID18751B08 dt:dt="string"></IVID18751B08>\n')
f.write('  <IVID86E1200 dt:dt="string"></IVID86E1200>\n')
f.write('  <IVID157115F8 dt:dt="string"></IVID157115F8>\n')
f.write('  <IVID1ACF422B dt:dt="string"></IVID1ACF422B>\n')
f.write('  <IVID406811FD dt:dt="string"></IVID406811FD>\n')
f.write('  <IVID376316F1 dt:dt="string"></IVID376316F1>\n')
f.write(' </CustomDocumentProperties>\n')
f.write(' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">\n')
f.write('  <WindowHeight>8445</WindowHeight>\n')
f.write('  <WindowWidth>7365</WindowWidth>\n')
f.write('  <WindowTopX>5985</WindowTopX>\n')
f.write('  <WindowTopY>-15</WindowTopY>\n')
f.write('  <TabRatio>844</TabRatio>\n')
f.write('  <ProtectStructure>False</ProtectStructure>\n')
f.write('  <ProtectWindows>False</ProtectWindows>\n')
f.write(' </ExcelWorkbook>\n')
f.write(' <Styles>\n')
f.write('  <Style ss:ID="Default" ss:Name="Normal">\n')
f.write('   <Alignment ss:Vertical="Bottom"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat/>\n')
f.write('   <Protection/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193440">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193460">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193480">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193500">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193520">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193540">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193560">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193580">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193600">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193620">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193216">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193236">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193256">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193276">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193296">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193316">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193336">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192992">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193012">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193032">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193052">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193072">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193092">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193112">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193132">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193152">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103193172">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192768">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192788">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192808">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192828">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192848">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192868">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192888">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192908">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192928">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192948">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192544">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192564">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192584">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192624">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="@"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192644">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m103192664">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="@"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s66">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s67">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s70">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s71">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s72">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s73">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s74">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s75">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s76">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s77">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s78">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s80">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s81">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s82">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s83">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s84">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s87">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s88">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s89">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s90">\n')
f.write('   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s91">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s95">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s109">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s121">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <NumberFormat/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s122">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="바탕" x:CharSet="129" x:Family="Roman" ss:Size="12"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s127">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <NumberFormat ss:Format="0.00000"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s130">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s131">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s137">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s139">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s141">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <NumberFormat ss:Format="0.00000_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s143">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s145">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s146">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s147">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <NumberFormat ss:Format="0.00000_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s209">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <NumberFormat ss:Format="0.0000_);[Red]\(0.0000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s210">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <NumberFormat ss:Format="0.000000_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s215">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s565">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="16"\n')
f.write('    ss:Bold="1"/>\n')
f.write('  </Style>\n')
f.write(' </Styles>\n')
f.write(' <Worksheet ss:Name="')
f.write('이화학 분석성적서')
f.write('">\n')
f.write('  <Table ss:ExpandedColumnCount="17" ss:ExpandedRowCount="186" x:FullColumns="1"\n')
f.write('   x:FullRows="1" ss:StyleID="s66" ss:DefaultColumnWidth="42"\n')
f.write('   ss:DefaultRowHeight="20.1">\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="4.5"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="72"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="51"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="36.75"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="32.25"\n')
f.write('    ss:Span="3"/>\n')
f.write('   <Column ss:Index="9" ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="28.5"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="29.25"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="24"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="27.75"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="29.25"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="5.25"/>\n')
f.write('   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="71.25"/>\n')
f.write('   <Column ss:StyleID="s66" ss:Width="63.75"/>\n')
f.write('   <Column ss:StyleID="s66" ss:Width="57"/>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="10.5">\n')
f.write('    <Cell ss:StyleID="s70"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s72"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeAcross="11" ss:StyleID="s565"><Data ss:Type="String">이화학 분석성적서</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="6.75">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="30">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s83"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">분 석 년 월 일</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="3" ss:StyleID="m103192624"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> ')
f.write(분석년월일)
f.write('</ss:Data></Cell>\n')
f.write(
    '    <Cell ss:MergeAcross="1" ss:StyleID="m103192644"><Data ss:Type="String">제 조 년 월 일&#10;(Batch No.)</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m103192664"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> ')
f.write(제조년)
f.write('년 ')
f.write(제조월)
f.write('월 ')
f.write(제조일)
f.write('일')
f.write('&#10;(')
f.write(LOTNO)
f.write(')</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="27">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s80"><Data ss:Type="String">분석  책임자</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s137"><Data ss:Type="String">소   속</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="3" ss:StyleID="m103192544"><Data ss:Type="String">')
f.write('(주)팜한농 작물보호연구센터')
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s137"><Data ss:Type="String">성  명</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s130"><Data ss:Type="String">  ')
f.write(책임자)
f.write(' (인)</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s130"/>\n')
f.write('    <Cell ss:StyleID="s130"/>\n')
f.write('    <Cell ss:StyleID="s130"/>\n')
f.write('    <Cell ss:StyleID="s146"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s91"><Data ss:Type="String">분석 의뢰자</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s87"><Data ss:Type="String">')
f.write('(주)팜한농 ')
f.write(의뢰자)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s95"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s88"/>\n')
f.write('    <Cell ss:StyleID="s89"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="15">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m103192564"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">품    목    명</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="10" ss:MergeDown="1" ss:StyleID="m103192584"><Data\n')
f.write('      ss:Type="String">')
f.write(str(한글명1))
if 한글명2 == "":
    pass
else:
    f.write('.')
    f.write(str(한글명2))

if 한글명3 == "":
    pass
else:
    f.write('.')
    f.write(str(한글명3))

f.write(' ')
f.write(str(제형분류))
f.write('&#10;(')
f.write(str(시료명1))

if 시료명2 == "":
    pass
else:
    f.write('.')
    f.write(str(시료명2))

if 시료명3 == "":
    pass
else:
    f.write('.')
    f.write(str(시료명3))

f.write(')</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="15">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="24">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="m103192768"><Data ss:Type="String">유효성분의 명칭&#10;및 함유량</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="10" ss:MergeDown="2" ss:StyleID="m103192788"><ss:Data\n')
f.write('      ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">')
f.write(시료명1IUPAC)
f.write('(IUPAC)……… ')
f.write(함량1)

if 시료명2IUPAC == "":
    pass
else:
    f.write('&#10; ')
    f.write(시료명2IUPAC)
    f.write('(IUPAC)……… ')
    f.write(함량2)

if 시료명3IUPAC == "":
    pass
else:
    f.write('&#10; ')
    f.write(시료명3IUPAC)
    f.write('(IUPAC)……… ')
    f.write(함량3)

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="24">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="24">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeAcross="11" ss:StyleID="m103192808"><Data ss:Type="String">분석결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="m103192828"><Data ss:Type="String">분석항목</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="m103192848"><Data ss:Type="String">분석회수</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="5" ss:StyleID="m103192868"><Data ss:Type="String">분      석      치</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="3" ss:MergeDown="2" ss:StyleID="m103192888"><Data\n')
f.write('      ss:Type="String">분석방법</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="18">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="m103192908"><Data\n')
f.write('      ss:Type="String">')
f.write(시료명1)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="m103192928"><Data\n')
f.write('      ss:Type="String">')
f.write(시료명2)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="m103192948"><Data\n')
f.write('      ss:Type="String">')
f.write(시료명3)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('    <Cell ss:Index="16" ss:StyleID="s122"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="21.75">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('    <Cell ss:Index="16" ss:StyleID="s122"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeDown="4" ss:StyleID="m103192992"><Data ss:Type="String"> 1. 유효성분</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s145"><Data ss:Type="Number">1</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193012"><Data ss:Type="String">')
f.write(str(sam_1_1_content))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193032"><Data ss:Type="String">')
f.write(str(sam_2_1_content))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193052"><Data ss:Type="String">')
f.write(str(sam_3_1_content))

f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="3" ss:MergeDown="2" ss:StyleID="m103193236"><Data\n')
f.write('      ss:Type="String">')
f.write(분석기기1)

if 분석기기2 == "":
    pass
else:
    f.write('&#10;')

f.write(분석기기2)

if 분석기기3 == "":
    pass
else:
    f.write('&#10;')

f.write(분석기기3)

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('    <Cell ss:Index="16" ss:StyleID="s209"/>\n')
f.write('    <Cell ss:StyleID="s209"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s131"><Data ss:Type="Number">2</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193132"><Data ss:Type="String">')
f.write(str(sam_1_2_content))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193152"><Data ss:Type="String">')
f.write(str(sam_2_2_content))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193256"><Data ss:Type="String">')
f.write(str(sam_3_2_content))

f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('    <Cell ss:Index="16" ss:StyleID="s210"/>\n')
f.write('    <Cell ss:StyleID="s210"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s215"><Data ss:Type="Number">3</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193172"><Data ss:Type="String">')
f.write(str(sam_1_3_content))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193316"><Data ss:Type="String">')
f.write(str(sam_2_3_content))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193336"><Data ss:Type="String">')
f.write(str(sam_3_3_content))

f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('    <Cell ss:Index="16" ss:StyleID="s147"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s215"><Data ss:Type="String">평 균 치</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193072"><Data ss:Type="String">')
f.write(str(sam_1_average))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193092"><Data ss:Type="String">')
f.write(str(sam_2_average))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193112"><Data ss:Type="String">')
f.write(str(sam_3_average))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="3" ss:StyleID="m103193540"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('    <Cell ss:StyleID="s121"/>\n')
f.write('    <Cell ss:StyleID="s127"/>\n')
f.write('    <Cell ss:StyleID="s127"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="46.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s131"><Data ss:Type="String">표준편차</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="9" ss:StyleID="m103193560"><Data ss:Type="String"> ')
f.write(시료명1)
f.write(' ± ')
f.write(str(sam_1_stdev))

if 시료명2 == "":
    pass
else:
    f.write('&#10; ')
    f.write(시료명2)
    f.write(' ± ')
    f.write(str(sam_2_stdev))

if 시료명3 == "":
    pass
else:
    f.write('&#10; ')
    f.write(시료명3)
    f.write(' ± ')
    f.write(str(sam_3_stdev))

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('    <Cell ss:StyleID="s121"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="37.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s139"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> 2. 물리성</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="10" ss:StyleID="m103193580"><Data ss:Type="String"> ')
f.write(검사항목1)

if 검사항목2 == "":
    pass
else:
    f.write('&#10; ')
    f.write(검사항목2)

if 검사항목3 == "":
    pass
else:
    f.write('&#10; ')
    f.write(검사항목3)

if 검사항목4 == "":
    pass
else:
    f.write('&#10; ')
    f.write(검사항목4)

if 검사항목5 == "":
    pass
else:
    f.write('&#10; ')
    f.write(검사항목5)

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s141"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="35.25">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s143"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> 3. 외 관</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s137"><Data ss:Type="String">성 상</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193276"><Data ss:Type="String">')
f.write(성상)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193296"><Data ss:Type="String">색상</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193600"><Data ss:Type="String">')
f.write(색상)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193620"><Data ss:Type="String">냄새</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m103193216"><Data ss:Type="String">')
f.write(냄새)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s141"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="35.25">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m103193460"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> 4. 시험항목</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="10" ss:MergeDown="1" ss:StyleID="m103193480"><Data\n')
f.write('      ss:Type="String"> ')
f.write(시험항목)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="8.25">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="26.25">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeAcross="11" ss:MergeDown="1" ss:StyleID="m103193500"><ss:Data\n')
f.write(
    '      ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">  첨부 자료&#10;&#10;        ○ 성적계산서, 분석법 및 HPLC chromatograms 첨부</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="26.25">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="19.5">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s109"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s71"/>\n')
f.write('    <Cell ss:StyleID="s72"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeAcross="11" ss:StyleID="m103193520"><Data ss:Type="String">')
f.write(분석년월일)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s81"/>\n')
f.write('    <Cell ss:StyleID="s74"/>\n')
f.write('    <Cell ss:StyleID="s74"/>\n')
f.write('    <Cell ss:StyleID="s74"/>\n')
f.write('    <Cell ss:StyleID="s74"/>\n')
f.write('    <Cell ss:StyleID="s74"/>\n')
f.write('    <Cell ss:StyleID="s74"/>\n')
f.write('    <Cell ss:StyleID="s90"/>\n')
f.write('    <Cell ss:StyleID="s90"/>\n')
f.write('    <Cell ss:StyleID="s90"/>\n')
f.write('    <Cell ss:StyleID="s90"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:MergeAcross="11" ss:StyleID="m103193440"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">주식회사 팜한농 작물보호연구센터장  ( 인)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.25">\n')
f.write('    <Cell ss:StyleID="s73"/>\n')
f.write('    <Cell ss:StyleID="s82"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s77"/>\n')
f.write('    <Cell ss:StyleID="s75"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="6.75">\n')
f.write('    <Cell ss:StyleID="s67"/>\n')
f.write('    <Cell ss:StyleID="s84"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s76"/>\n')
f.write('    <Cell ss:StyleID="s77"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s78"/>\n')
f.write('   </Row>\n')
f.write('  </Table>\n')
f.write('  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n')
f.write('   <PageSetup>\n')
f.write('    <Layout x:CenterHorizontal="1" x:CenterVertical="1"/>\n')
f.write('    <Header x:Margin="0"/>\n')
f.write('    <Footer x:Margin="0"/>\n')
f.write('    <PageMargins x:Bottom="0" x:Left="0" x:Right="0" x:Top="0"/>\n')
f.write('   </PageSetup>\n')
f.write('   <Unsynced/>\n')
f.write('   <Print>\n')
f.write('    <ValidPrinterInfo/>\n')
f.write('    <PaperSizeIndex>9</PaperSizeIndex>\n')
f.write('    <HorizontalResolution>600</HorizontalResolution>\n')
f.write('    <VerticalResolution>600</VerticalResolution>\n')
f.write('   </Print>\n')
f.write('   <Selected/>\n')
f.write('   <Panes>\n')
f.write('    <Pane>\n')
f.write('     <Number>3</Number>\n')
f.write('     <ActiveRow>24</ActiveRow>\n')
f.write('     <ActiveCol>14</ActiveCol>\n')
f.write('    </Pane>\n')
f.write('   </Panes>\n')
f.write('   <ProtectObjects>False</ProtectObjects>\n')
f.write('   <ProtectScenarios>False</ProtectScenarios>\n')
f.write('  </WorksheetOptions>\n')
f.write(' </Worksheet>\n')
f.write('</Workbook>\n')

f.close()
