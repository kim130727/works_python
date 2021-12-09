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

pypyodbc.lowercase = False
conn = pypyodbc.connect(

    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\data automation\이화학적분석.accdb;")

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
시험번호 = r[1]
시험의뢰자 = r[6]
시험책임자 = r[7]
시험담당자 = r[8]
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
검사항목1결과 = ""
검사항목2결과 = ""
검사항목3결과 = ""
검사항목4결과 = ""
검사항목5결과 = ""

if 제형코드 == "EC":
    제형분류 = "유제"
    수량 = "500ml"
    검사항목1 = "유화성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[77]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE 재질"

elif 제형코드 == "SL":
    제형분류 = "액제"
    검사항목1 = "수용성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[80]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE 재질"

    수량 = "500ml"

elif 제형코드 == "SC":
    제형분류 = "액상수화제"
    수량 = "500ml"
    검사항목1 = "수화성"
    검사항목2 = "분말도"
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[78]
    검사항목2결과 = r[81]
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE 재질"

elif 제형코드 == "WP":
    제형분류 = "수화제"
    수량 = "500g"
    검사항목1 = "수화성"
    검사항목2 = "분말도"
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[78]
    검사항목2결과 = r[81]
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "은박코팅봉투, 알루미늄재질"
elif 제형코드 == "SP":
    제형분류 = "수용제"
    수량 = "500g"
elif 제형코드 == "OD":
    제형분류 = "유상수화제"
    수량 = "500ml"
    검사항목1 = "유화성"
    검사항목2 = "수화성"
    검사항목3 = "분말도"
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[77]
    검사항목2결과 = r[78]
    검사항목3결과 = r[81]
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE재질"
elif 제형코드 == "EP":
    제형분류 = "분상유제"
    수량 = "500g"
elif 제형코드 == "DP":
    제형분류 = "분제"
    수량 = "500g"
    검사항목1 = "분말도"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[81]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
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
    검사항목1 = ""
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = ""
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "은박코팅봉투, 알루미늄재질"
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
    검사항목1 = "발연성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[94]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
elif 제형코드 == "WF":
    제형분류 = "수화성미분제"
    수량 = "500g"
elif 제형코드 == "WG":
    제형분류 = "입상수화제"
    수량 = "500g"
    검사항목1 = "수화성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[78]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, PE재질"
elif 제형코드 == "EW":
    제형분류 = "유탁제"
    수량 = "500ml"
    검사항목1 ="유화성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[77]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
elif 제형코드 == "CS":
    제형분류 = "캡슐현탁제"
    수량 = "500ml"
elif 제형코드 == "SE":
    제형분류 = "유현탁제"
    수량 = "500ml"
    검사항목1 = "유화성"
    검사항목2 = "수화성"
    검사항목3 = "분말도"
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[77]
    검사항목2결과 = r[78]
    검사항목3결과 = r[81]
    검사항목4결과 = ""
    검사항목5결과 = ""
elif 제형코드 == "DC":
    제형분류 = "분산성액제"
    수량 = "500ml"
    검사항목1 = "수중분산성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[79]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
elif 제형코드 == "SO":
    제형분류 = "수면전개제"
    수량 = "500ml"
    검사항목1 = ""
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = ""
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE 재질"
elif 제형코드 == "WS":
    제형분류 = "종자처리수화제"
    수량 = "500g"
    검사항목1 = "수화성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[78]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "은박코팅봉투, 알루미늄재질"
elif 제형코드 == "ME":
    제형분류 = "미탁제"
    수량 = "500ml"
elif 제형코드 == "FS":
    제형분류 = "종자처리액상수화제"
    수량 = "500ml"
    검사항목1 = "수화성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[78]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE재질"
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
    포장용기 = "합성수지병, PE재질"
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
    검사항목1 = "확산성"
    검사항목2 = ""
    검사항목3 = ""
    검사항목4 = ""
    검사항목5 = ""
    검사항목1결과 = r[93]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "은박코팅봉투, 알루미늄재질"
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

print ("********************************************************************************************")
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

print("********************************************************************************************")
print("농약품목의 이화학적 분석 보고서")
print(한글명1, 한글명2, 한글명3, 제형분류, " 이화학적 분석")
print("시험번호 : ", 시험번호)
print("(주)팜한농 작물보호연구센터")
print("********************************************************************************************")
print("제출문")
print("제목 :",시험물질,"의 이화학적 분석")
print("시험번호 : ", 시험번호)
print("본 시험에 사용된 기준은 다음과 같습니다.")
print("1. 농촌진흥청 고시 농약 및 원제의 등록기준 및 농약의 검사방법 및 부정불량 농약 처리요령")
print("2. 농촌진흥청 고시 농약 등의 시험연구기고나 지정 및 관리기준")
print("본 보고서에 기술된 시험과정은 시험책임자의 책임 하에 수행되었으며,")
print("위의 기준을 준수하여 실시하였으며, 시험결과는 생성된 모든 시험기초자료를 토대로 작성되었습니다.")
print("시험 책임자 : ",시험책임자)
print("********************************************************************************************")
print("농약품목의 이화학적 분석성적서")
print("시험번호 ", 시험번호, ", 시험분야 : 이화학, 시험년도 :",분석일[0:4])
print("시험항목 ",시험물질,"의 이화학적 분석")
print("시험기간 ", 분석년월일)
print("시험기관: (주)팜한농 작물보호연구센터")
print("시험담당자: ", 시험담당자)
print("시험의뢰자: ", 시험의뢰자)
print("")
print("1. 목적")
print (시험물질,"에 대한 이화학성을 구명하여 농약품목등록의 이화학성 평가 및 품질관리 기준설정")
print ("등을 위한 기초자료로 활용코자 함")
print ("2. 시험방법")
print ("가. 시험약제 :", 시험물질)
print ("시험물질정보 :", LOTNO, '제조, 제조일자:', 제조년 + '년', 제조월 + '월', 제조일 + '일')
print ("나. 시험세부항목")
print ("유효성분 함량(%)")
print ("품목의 물리성 :", 검사항목1, 검사항목2, 검사항목3, 검사항목4, 검사항목5)
print ("품목의 외관 : 성상, 색, 냄새, 포장용기 및 재질")
print ("")
print ("3. 시험성적")
print (시험물질, " 품목의 유효성분")
print ("유효성분 :", 시료명1)
print ("표준품 순도(%) ", std_content1, "표준품 무게 ",std_g1,"시료 무게 ",sam_1_1_g," " ,sam_1_2_g," ",sam_1_3_g)
print('유효성분 함량:', sam_1_1_content, sam_1_2_content, sam_1_3_content, '평균', sam_1_average)
print ("유효성분 :", 시료명2)
print ("표준품 순도(%) ", std_content2, "표준품 무게 ",std_g2,"시료 무게 ",sam_2_1_g," " ,sam_2_2_g," ",sam_2_3_g)
print('유효성분 함량:', sam_2_1_content, sam_2_2_content, sam_2_3_content, '평균', sam_2_average)
print ("유효성분 :", 시료명3)
print ("표준품 순도(%) ", std_content3, "표준품 무게 ",std_g3,"시료 무게 ",sam_3_1_g," " ,sam_3_2_g," ",sam_3_3_g)
print('유효성분 함량:', sam_3_1_content, sam_3_2_content, sam_3_3_content, '평균', sam_3_average)
print("")
print("품목의 물리성")
print(검사항목1," : ",검사항목1결과)
print(검사항목2," : ",검사항목2결과)
print(검사항목3," : ",검사항목3결과)
print(검사항목4," : ",검사항목4결과)
print(검사항목5," : ",검사항목5결과)
print ("")
print("품목의 외관등")
print("성상: ",성상, ",색상: ", 색상, ",냄새: ", 냄새, ",포장용기 및 재질: ", 포장용기)
print("")
print("4. 시험결과 요약")
print("가. 이화학분석시험결과, 유효성분의 함량은", 시료명1, sam_1_average, 시료명2, sam_2_average, 시료명3, sam_3_average,"로 등록기준에 적합하였다.")
print("나.",검사항목1, 검사항목2, 검사항목3, 검사항목4, 검사항목5,"는 양호하였다.")
print("다. 이상의 결과를 종합하여",시험물질,"는 농약품목 등록에 이화학적으로 적합한 약제로서")
print("품질관리 기준설정에 문제가 없는 것으로 판단되었다.")
print("")
print("5.첨부자료")
print("가. 유효성분 분석 성적계산서 및 크로마토그램")