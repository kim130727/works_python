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

nm = input('시험번호를 입력해 주세요 ->')

pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\data automation\이화학적분석.accdb;")

ope1 = '시작'
ope2 = '2주'
ope3 = '4주'
ope4 = '6주'
ope5 = '8주'
ope6 = '저온'

c1 = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c1.execute(sql, (nm, ope1))
r1 = c1.fetchone()

c2 = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c2.execute(sql, (nm, ope2))
r2 = c2.fetchone()

c3 = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c3.execute(sql, (nm, ope3))
r3 = c3.fetchone()

c4 = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c4.execute(sql, (nm, ope4))
r4 = c4.fetchone()

c5 = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c5.execute(sql, (nm, ope5))
r5 = c5.fetchone()

c6 = conn.cursor()
sql = 'SELECT * FROM Total WHERE 시험번호=? and 시험내용=?'
c6.execute(sql, (nm, ope6))
r6 = c6.fetchone()

print(r1)
print(r2)
print(r3)
print(r4)
print(r5)
print(r6)

구분_경변시작 = r1[2]
시험내용_경변시작 = r1[10]
시료명1_경변시작 = r1[18]
한글명1_경변시작 = r1[19]
시료명2_경변시작 = r1[38]
한글명2_경변시작 = r1[39]


if 한글명2_경변시작 == None:
    한글명2_경변시작 = ""

시료명3_경변시작 = r1[58]
한글명3_경변시작 = r1[59]

if 한글명3_경변시작 == None:
    한글명3_경변시작 = ""

시험물질_경변시작 = r1[3]
분석일_경변시작 = r1[9]
분석년월일_경변시작 = list(re.findall(r"(\d+)", 분석일_경변시작))[0] + "년" + list(re.findall(r"(\d+)", 분석일_경변시작))[1] + "월" + \
             list(re.findall(r"(\d+)", 분석일_경변시작))[2] + "일"
분석기기1_경변시작 = r1[11]
분석기기2_경변시작 = r1[12]
분석기기3_경변시작 = r1[13]
성상_경변시작 = r1[74]
색상_경변시작 = r1[75]
냄새_경변시작 = r1[76]
책임자 = r1[7]
의뢰자 = r1[6]
성상_경변시작 = r1[74]
색상_경변시작 = r1[76]

if 성상_경변시작 == None:
    성상_경변시작 = ""
else:
    pass

if 색상_경변시작 == None:
    색상_경변시작 = ""
else:
    pass

if 냄새_경변시작 == None:
    냄새_경변시작 = ""
else:
    pass

if 분석기기1_경변시작 == 분석기기2_경변시작:
    분석기기2_경변시작 = ""
else:
    pass

if 분석기기1_경변시작 == 분석기기3_경변시작:
    분석기기3_경변시작 = ""
else:
    pass

if 분석기기2_경변시작 == None:
    분석기기2_경변시작 = ""
else:
    pass

if 분석기기3_경변시작 == None:
    분석기기3_경변시작 = ""
else:
    pass

try:
    if 시료명1_경변시작 == None:
        시료명1IUPAC_경변시작 = ""
        시료명1_경변시작 = ""
    else:
        시료명1CID_경변시작 = pcp.get_compounds(시료명1_경변시작, 'name')
        시료명1IUPAC_경변시작 = 시료명1CID_경변시작[0].iupac_name
except IndexError:
    시료명1IUPAC_경변시작 = ""


try:
    if 시료명2_경변시작 == None:
        시료명2IUPAC_경변시작 = ""
        시료명2_경변시작 = ""
    else:
        시료명2CID_경변시작 = pcp.get_compounds(시료명2_경변시작, 'name')
        시료명2IUPAC_경변시작 = 시료명2CID_경변시작[0].iupac_name
except IndexError:
    시료명2IUPAC_경변시작 = ""

try:
    if 시료명3_경변시작 == None:
        시료명3IUPAC_경변시작 = ""
        시료명3_경변시작 = ""
    else:
        시료명3CID_경변시작 = pcp.get_compounds(시료명3_경변시작, 'name')
        시료명3IUPAC_경변시작 = 시료명3CID_경변시작[0].iupac_name
except IndexError:
    시료명3IUPAC_경변시작 = ""

제형코드_경변시작 = 시험물질_경변시작[-2:]
LOTNO_경변시작 = r1[4]
제조년_경변시작 = '20' + list(re.findall(r"(\d+)", LOTNO_경변시작)[0])[0] + list(re.findall(r"(\d+)", LOTNO_경변시작)[0])[1]
제조월_경변시작 = list(re.findall(r"(\d+)", LOTNO_경변시작)[0])[2] + list(re.findall(r"(\d+)", LOTNO_경변시작)[0])[3]
제조일_경변시작 = list(re.findall(r"(\d+)", LOTNO_경변시작)[0])[4] + list(re.findall(r"(\d+)", LOTNO_경변시작)[0])[5]

검사항목1_경변시작 = ""
검사항목2_경변시작 = ""
검사항목3_경변시작 = ""
검사항목4_경변시작 = ""
검사항목5_경변시작 = ""

if 제형코드_경변시작 == "EC":
    제형분류_경변시작 = "유제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    검사항목1_경변시작 = "유화성: " + r1[77]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "SL":
    제형분류_경변시작 = "액제"
    수량_경변시작 = "200ml"
    검사항목1_경변시작 = "수용성: " + r1[80]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
    검사항목1결과 = r1[80]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
elif 제형코드_경변시작 == "SC":
    제형분류_경변시작 = "액상수화제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    검사항목1_경변시작 = "수화성: " + r1[78]
    검사항목2_경변시작 = "분말도: " + r1[81]
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    포장용기단위_경변시작 = "병"
    검사항목1결과 = r1[78]
    검사항목2결과 = r1[81]
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "합성수지병, HDPE 재질"
elif 제형코드_경변시작 == "WP":
    제형분류_경변시작 = "수화제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "SP":
    제형분류_경변시작 = "수용제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "OD":
    제형분류_경변시작 = "유상수화제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "EP":
    제형분류_경변시작 = "분상유제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "DP":
    제형분류_경변시작 = "분제"
    수량_경변시작 = "500g"
    검사항목1_경변시작 = "분말도: " + r1[81]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "DS":
    제형분류_경변시작 = "분의제"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
    수량_경변시작 = "500g"
elif 제형코드_경변시작 == "GP":
    제형분류_경변시작 = "미분제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "DL":
    제형분류_경변시작 = "저비산분제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"

elif 제형코드_경변시작 == "GR":
    제형분류_경변시작 = "입제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "MG":
    제형분류_경변시작 = "미립제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "PA":
    제형분류_경변시작 = "도포제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "GA":
    제형분류_경변시작 = "훈증제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "FU":
    제형분류_경변시작 = "훈연제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "AE":
    제형분류_경변시작 = "연무제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "CG":
    제형분류_경변시작 = "캡슐제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
elif 제형코드_경변시작 == "FG":
    제형분류_경변시작 = "세립제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "FW":
    제형분류_경변시작 = "과립훈연제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
    검사항목1_경변시작 = "발연성: " + r1[94]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
elif 제형코드_경변시작 == "WF":
    제형분류_경변시작 = "수화성미분제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "WG":
    제형분류_경변시작 = "입상수화제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "합성수지병, PE재질"
    검사항목1_경변시작 = "수화성: " + r1[78]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
elif 제형코드_경변시작 == "EW":
    제형분류_경변시작 = "유탁제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
    검사항목1_경변시작 = "유화성: " + r1[77]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
elif 제형코드_경변시작 == "CS":
    제형분류_경변시작 = "캡슐현탁제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "SE":
    제형분류_경변시작 = "유현탁제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
    검사항목1_경변시작 = "유화성: " + r1[77]
    검사항목2_경변시작 = "수화성: " + r1[78]
    검사항목3_경변시작 = "분말도: " + r1[81]
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
elif 제형코드_경변시작 == "DC":
    제형분류_경변시작 = "분산성액제"
    수량_경변시작 = "500ml"
    검사항목1_경변시작 = "수중분산성: " + r1[79]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "SO":
    제형분류_경변시작 = "수면전개제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "WS":
    제형분류_경변시작 = "종자처리수화제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "ME":
    제형분류_경변시작 = "미탁제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "FS":
    제형분류_경변시작 = "종자처리액상수화제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "UG":
    제형분류_경변시작 = "수면부상성입제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "수용성필름, 은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "PF":
    제형분류_경변시작 = "비닐멀칭제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "SF":
    제형분류_경변시작 = "판상줄제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "OL":
    제형분류_경변시작 = "오일제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "SG":
    제형분류_경변시작 = "입상수용제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "AL":
    제형분류_경변시작 = "직접살포액제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
elif 제형코드_경변시작 == "DT":
    제형분류_경변시작 = "직접살포정제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
    포장용기단위_경변시작 = "봉"
elif 제형코드_경변시작 == "VP":
    제형분류_경변시작 = "마이크로캡슐훈증제"
    수량_경변시작 = ""
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "AS":
    제형분류_경변시작 = "액상제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "EM":
    제형분류_경변시작 = "유상현탁제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "SM":
    제형분류_경변시작 = "액상현탁제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "GM":
    제형분류_경변시작 = "고상제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "GG":
    제형분류_경변시작 = "대립제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
    검사항목1_경변시작 = "확산성: " + r1[93]
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
elif 제형코드_경변시작 == "WT":
    제형분류_경변시작 = "정제상수화제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "ZC":
    제형분류_경변시작 = "캡슐액상수화제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
elif 제형코드_경변시작 == "GD":
    제형분류_경변시작 = "발생기"
    수량_경변시작 = ""
    포장용기_경변시작 = ""
elif 제형코드_경변시작 == "RM":
    제형분류_경변시작 = "발생제"
    수량_경변시작 = ""
    포장용기_경변시작 = ""
else:
    제형분류_경변시작 = ""
    수량_경변시작 = ""
    포장용기_경변시작 = ""
    포장용기단위_경변시작 = "봉"

# print(시험물질_경변시작)

시험물질_pre_경변시작 = re.sub(r"([ (+)%])+", "             ", 시험물질_경변시작)

# print('시험물질pre', 시험물질_pre_경변시작)
함량_경변시작 = re.findall(r"(\d\d?.?\d?\d?\d?)", 시험물질_pre_경변시작)

함량_경변시작.extend(["", "", ""])
# print('함량', 함량_경변시작)

함량1_경변시작 = 함량_경변시작[0] + '%'

if 함량_경변시작[1] == "":
    함량2_경변시작 = ""
else:
    함량2_경변시작 = 함량_경변시작[1] + '%'

if 함량_경변시작[2] == "":
    함량3_경변시작 = ""
else:
    함량3_경변시작 = 함량_경변시작[2] + '%'

소수점1_경변시작 = re.findall(r"([.]\d+)", 함량_경변시작[0])
소수점2_경변시작 = re.findall(r"([.]\d+)", 함량_경변시작[1])
소수점3_경변시작 = re.findall(r"([.]\d+)", 함량_경변시작[2])

if len(str(소수점1_경변시작)) == 2:
    point1 = 2
else:
    point1 = len(str(소수점1_경변시작)) - 3

if len(str(소수점2_경변시작)) == 2:
    point2 = 2
else:
    point2 = len(str(소수점2_경변시작)) - 3

if len(str(소수점3_경변시작)) == 2:
    point3 = 2
else:
    point3 = len(str(소수점3_경변시작)) - 3

# print('소수점1', 소수점1_경변시작, len(str(소수점1_경변시작)))
# print('소수점2', 소수점2_경변시작, len(str(소수점2_경변시작)))
# print('소수점3', 소수점3_경변시작, len(str(소수점3_경변시작)))
# print('point1', point1)
# print('point2', point2)
# print('point3', point3)

std_content1_경변시작 = r1[14]
std_g1_경변시작 = r1[15]
std_AI_area1_경변시작 = r1[16]
std_IS_area1_경변시작 = r1[17]

sam_1_1_g_경변시작 = r1[20]
sam_1_1_AI_경변시작 = r1[21]
sam_1_1_IS_경변시작 = r1[22]

sam_1_2_g_경변시작 = r1[24]
sam_1_2_AI_경변시작 = r1[25]
sam_1_2_IS_경변시작 = r1[26]

sam_1_3_g_경변시작 = r1[28]
sam_1_3_AI_경변시작 = r1[29]
sam_1_3_IS_경변시작 = r1[30]

factor_std1_경변시작 = round(
    float(std_g1_경변시작) * float(std_content1_경변시작) * float(std_IS_area1_경변시작) / float(std_AI_area1_경변시작), 4)
sam_1_1_content_경변시작 = round(
    (factor_std1_경변시작 * float(sam_1_1_AI_경변시작)) / (float(sam_1_1_IS_경변시작) * (float(sam_1_1_g_경변시작))),
    int(point1))
sam_1_2_content_경변시작 = round(
    (factor_std1_경변시작 * float(sam_1_2_AI_경변시작)) / (float(sam_1_2_IS_경변시작) * (float(sam_1_2_g_경변시작))),
    int(point1))
sam_1_3_content_경변시작 = round(
    (factor_std1_경변시작 * float(sam_1_3_AI_경변시작)) / (float(sam_1_3_IS_경변시작) * (float(sam_1_3_g_경변시작))),
    int(point1))
sam_1_average_경변시작 = round((sam_1_1_content_경변시작 + sam_1_2_content_경변시작 + sam_1_3_content_경변시작) / 3,
                           int(point1))
sam_1_stdev_경변시작 = round(((((sam_1_1_content_경변시작 - (
    (sam_1_1_content_경변시작 + sam_1_2_content_경변시작 + sam_1_3_content_경변시작) / 3)) ** 2 + (sam_1_2_content_경변시작 - (
    (sam_1_1_content_경변시작 + sam_1_2_content_경변시작 + sam_1_3_content_경변시작) / 3)) ** 2 + (sam_1_3_content_경변시작 - (
    (sam_1_1_content_경변시작 + sam_1_2_content_경변시작 + sam_1_3_content_경변시작) / 3)) ** 2)) / 2) ** 0.5, 5)

if 시료명2_경변시작 == "":
    factor_std2_경변시작 = ""
    sam_2_1_content_경변시작 = ""
    sam_2_2_content_경변시작 = ""
    sam_2_3_content_경변시작 = ""
    sam_2_average_경변시작 = ""
    sam_2_stdev_경변시작 = ""
    한글명2_경변시작 = ""

else:
    std_content2_경변시작 = r1[34]
    std_g2_경변시작 = r1[35]
    std_AI_area2_경변시작 = r1[36]
    std_IS_area2_경변시작 = r1[37]

    sam_2_1_g_경변시작 = r1[40]
    sam_2_1_AI_경변시작 = r1[41]
    sam_2_1_IS_경변시작 = r1[42]

    sam_2_2_g_경변시작 = r1[44]
    sam_2_2_AI_경변시작 = r1[45]
    sam_2_2_IS_경변시작 = r1[46]

    sam_2_3_g_경변시작 = r1[48]
    sam_2_3_AI_경변시작 = r1[49]
    sam_2_3_IS_경변시작 = r1[50]

    factor_std2_경변시작 = round(
        float(std_g2_경변시작) * float(std_content2_경변시작) * float(std_IS_area2_경변시작) / float(std_AI_area2_경변시작), 4)
    sam_2_1_content_경변시작 = round(
        (factor_std2_경변시작 * float(sam_2_1_AI_경변시작)) / (float(sam_2_1_IS_경변시작) * (float(sam_2_1_g_경변시작))), int(point2))
    sam_2_2_content_경변시작 = round(
        (factor_std2_경변시작 * float(sam_2_2_AI_경변시작)) / (float(sam_2_2_IS_경변시작) * (float(sam_2_2_g_경변시작))), int(point2))
    sam_2_3_content_경변시작 = round(
        (factor_std2_경변시작 * float(sam_2_3_AI_경변시작)) / (float(sam_2_3_IS_경변시작) * (float(sam_2_3_g_경변시작))), int(point2))
    sam_2_average_경변시작 = round((sam_2_1_content_경변시작 + sam_2_2_content_경변시작 + sam_2_3_content_경변시작) / 3, int(point2))
    sam_2_stdev_경변시작 = round(((((sam_2_1_content_경변시작 - (
    (sam_2_1_content_경변시작 + sam_2_2_content_경변시작 + sam_2_3_content_경변시작) / 3)) ** 2 + (sam_2_2_content_경변시작 - (
    (sam_2_1_content_경변시작 + sam_2_2_content_경변시작 + sam_2_3_content_경변시작) / 3)) ** 2 + (sam_2_3_content_경변시작 - (
    (sam_2_1_content_경변시작 + sam_2_2_content_경변시작 + sam_2_3_content_경변시작) / 3)) ** 2)) / 2) ** 0.5, 5)

if 시료명3_경변시작 == "":
    factor_std3_경변시작 = ""
    sam_3_1_content_경변시작 = ""
    sam_3_2_content_경변시작 = ""
    sam_3_3_content_경변시작 = ""
    sam_3_average_경변시작 = ""
    sam_3_stdev_경변시작 = ""
    한글명3_경변시작 = ""

else:
    std_content3_경변시작 = r1[54]
    std_g3_경변시작 = r1[55]
    std_AI_area3_경변시작 = r1[56]
    std_IS_area3_경변시작 = r1[57]

    sam_3_1_g_경변시작 = r1[60]
    sam_3_1_AI_경변시작 = r1[61]
    sam_3_1_IS_경변시작 = r1[62]

    sam_3_2_g_경변시작 = r1[64]
    sam_3_2_AI_경변시작 = r1[65]
    sam_3_2_IS_경변시작 = r1[66]

    sam_3_3_g_경변시작 = r1[68]
    sam_3_3_AI_경변시작 = r1[69]
    sam_3_3_IS_경변시작 = r1[70]

    factor_std3_경변시작 = round(
        float(std_g3_경변시작) * float(std_content3_경변시작) * float(std_IS_area3_경변시작) / float(std_AI_area3_경변시작), 4)
    sam_3_1_content_경변시작 = round(
        (factor_std3_경변시작 * float(sam_3_1_AI_경변시작)) / (float(sam_3_1_IS_경변시작) * (float(sam_3_1_g_경변시작))), int(point3))
    sam_3_2_content_경변시작 = round(
        (factor_std3_경변시작 * float(sam_3_2_AI_경변시작)) / (float(sam_3_2_IS_경변시작) * (float(sam_3_2_g_경변시작))), int(point3))
    sam_3_3_content_경변시작 = round(
        (factor_std3_경변시작 * float(sam_3_3_AI_경변시작)) / (float(sam_3_3_IS_경변시작) * (float(sam_3_3_g_경변시작))), int(point3))
    sam_3_average_경변시작 = round((sam_3_1_content_경변시작 + sam_3_2_content_경변시작 + sam_3_3_content_경변시작) / 3, int(point3))
    sam_3_stdev_경변시작 = round(((((sam_3_1_content_경변시작 - (
    (sam_3_1_content_경변시작 + sam_3_2_content_경변시작 + sam_3_3_content_경변시작) / 3)) ** 2 + (sam_3_2_content_경변시작 - (
    (sam_3_1_content_경변시작 + sam_3_2_content_경변시작 + sam_3_3_content_경변시작) / 3)) ** 2 + (sam_3_3_content_경변시작 - (
    (sam_3_1_content_경변시작 + sam_3_2_content_경변시작 + sam_3_3_content_경변시작) / 3)) ** 2)) / 2) ** 0.5, 5)

if r2 == None:
    분석일_경변1차 = ""
    분석년월일_경변1차 = ""
    시료명1_경변1차 = ""
    시료명2_경변1차 = ""
    시료명3_경변1차 = ""
    factor_std1_경변1차 = ""
    sam_1_1_content_경변1차 = ""
    sam_1_2_content_경변1차 = ""
    sam_1_3_content_경변1차 = ""
    sam_1_average_경변1차 = ""
    sam_1_stdev_경변1차 = ""
    시료1경변분해율1차 = ""
    factor_std2_경변1차 = ""
    sam_2_1_content_경변1차 = ""
    sam_2_2_content_경변1차 = ""
    sam_2_3_content_경변1차 = ""
    sam_2_average_경변1차 = ""
    sam_2_stdev_경변1차 = ""
    시료2경변분해율1차 = ""
    factor_std3_경변1차 = ""
    sam_3_1_content_경변1차 = ""
    sam_3_2_content_경변1차 = ""
    sam_3_3_content_경변1차 = ""
    sam_3_average_경변1차 = ""
    sam_3_stdev_경변1차 = ""
    시료3경변분해율1차 = ""

else:
    시료명1_경변1차 = r2[18]
    시료명2_경변1차 = r2[38]
    시료명3_경변1차 = r2[58]
    분석일_경변1차 = r2[9]
    분석년월일_경변1차 = list(re.findall(r"(\d+)", 분석일_경변1차))[0] + "년" + list(re.findall(r"(\d+)", 분석일_경변1차))[1] + "월" + \
                 list(re.findall(r"(\d+)", 분석일_경변1차))[2] + "일"

    try:
        if 시료명1_경변1차 == None:
            시료명1IUPAC_경변1차 = ""
            시료명1_경변1차 = ""
        else:
            시료명1CID_경변1차 = pcp.get_compounds(시료명1_경변1차, 'name')
            시료명1IUPAC_경변1차 = 시료명1CID_경변1차[0].iupac_name
    except IndexError:
        시료명1IUPAC_경변1차 = ""


    try:
        if 시료명2_경변1차 == None:
            시료명2IUPAC_경변1차 = ""
            시료명2_경변1차 = ""
        else:
            시료명2CID_경변1차 = pcp.get_compounds(시료명2_경변1차, 'name')
            시료명2IUPAC_경변1차 = 시료명2CID_경변1차[0].iupac_name
    except IndexError:
        시료명2IUPAC_경변1차 = ""


    try:
        if 시료명3_경변1차 == None:
            시료명3IUPAC_경변1차 = ""
            시료명3_경변1차 = ""
        else:
            시료명3CID_경변1차 = pcp.get_compounds(시료명3_경변1차, 'name')
            시료명3IUPAC_경변1차 = 시료명3CID_경변1차[0].iupac_name
    except IndexError:
        시료명3IUPAC_경변1차 = ""


    std_content1_경변1차 = r2[14]
    std_g1_경변1차 = r2[15]
    std_AI_area1_경변1차 = r2[16]
    std_IS_area1_경변1차 = r2[17]

    sam_1_1_g_경변1차 = r2[20]
    sam_1_1_AI_경변1차 = r2[21]
    sam_1_1_IS_경변1차 = r2[22]

    sam_1_2_g_경변1차 = r2[24]
    sam_1_2_AI_경변1차 = r2[25]
    sam_1_2_IS_경변1차 = r2[26]

    sam_1_3_g_경변1차 = r2[28]
    sam_1_3_AI_경변1차 = r2[29]
    sam_1_3_IS_경변1차 = r2[30]

    factor_std1_경변1차 = round(
        float(std_g1_경변1차) * float(std_content1_경변1차) * float(std_IS_area1_경변1차) / float(std_AI_area1_경변1차), 4)
    sam_1_1_content_경변1차 = round(
        (factor_std1_경변1차 * float(sam_1_1_AI_경변1차)) / (float(sam_1_1_IS_경변1차) * (float(sam_1_1_g_경변1차))),
        int(point1))
    sam_1_2_content_경변1차 = round(
        (factor_std1_경변1차 * float(sam_1_2_AI_경변1차)) / (float(sam_1_2_IS_경변1차) * (float(sam_1_2_g_경변1차))),
        int(point1))
    sam_1_3_content_경변1차 = round(
        (factor_std1_경변1차 * float(sam_1_3_AI_경변1차)) / (float(sam_1_3_IS_경변1차) * (float(sam_1_3_g_경변1차))),
        int(point1))
    sam_1_average_경변1차 = round((sam_1_1_content_경변1차 + sam_1_2_content_경변1차 + sam_1_3_content_경변1차) / 3,
                               int(point1))
    sam_1_stdev_경변1차 = round(((((sam_1_1_content_경변1차 - (
        (sam_1_1_content_경변1차 + sam_1_2_content_경변1차 + sam_1_3_content_경변1차) / 3)) ** 2 + (
                                    sam_1_2_content_경변1차 - ((
                                                                sam_1_1_content_경변1차 + sam_1_2_content_경변1차 + sam_1_3_content_경변1차) / 3)) ** 2 + (
                                    sam_1_3_content_경변1차 - ((
                                                                sam_1_1_content_경변1차 + sam_1_2_content_경변1차 + sam_1_3_content_경변1차) / 3)) ** 2)) / 2) ** 0.5,
                             5)

    시료1경변분해율1차 = round(
        ((float(sam_1_average_경변시작) - float(sam_1_average_경변1차)) / float(sam_1_average_경변시작)) * 100, 2)

    if 시료명2_경변1차 == "":
        factor_std2_경변1차 = ""
        sam_2_1_content_경변1차 = ""
        sam_2_2_content_경변1차 = ""
        sam_2_3_content_경변1차 = ""
        sam_2_average_경변1차 = ""
        sam_2_stdev_경변1차 = ""
        시료2경변분해율1차 = ""

    else:
        std_content2_경변1차 = r2[34]
        std_g2_경변1차 = r2[35]
        std_AI_area2_경변1차 = r2[36]
        std_IS_area2_경변1차 = r2[37]

        sam_2_1_g_경변1차 = r2[40]
        sam_2_1_AI_경변1차 = r2[41]
        sam_2_1_IS_경변1차 = r2[42]

        sam_2_2_g_경변1차 = r2[44]
        sam_2_2_AI_경변1차 = r2[45]
        sam_2_2_IS_경변1차 = r2[46]

        sam_2_3_g_경변1차 = r2[48]
        sam_2_3_AI_경변1차 = r2[49]
        sam_2_3_IS_경변1차 = r2[50]

        factor_std2_경변1차 = round(
            float(std_g2_경변1차) * float(std_content2_경변1차) * float(std_IS_area2_경변1차) / float(std_AI_area2_경변1차), 4)
        sam_2_1_content_경변1차 = round(
            (factor_std2_경변1차 * float(sam_2_1_AI_경변1차)) / (float(sam_2_1_IS_경변1차) * (float(sam_2_1_g_경변1차))),
            int(point2))
        sam_2_2_content_경변1차 = round(
            (factor_std2_경변1차 * float(sam_2_2_AI_경변1차)) / (float(sam_2_2_IS_경변1차) * (float(sam_2_2_g_경변1차))),
            int(point2))
        sam_2_3_content_경변1차 = round(
            (factor_std2_경변1차 * float(sam_2_3_AI_경변1차)) / (float(sam_2_3_IS_경변1차) * (float(sam_2_3_g_경변1차))),
            int(point2))
        sam_2_average_경변1차 = round((sam_2_1_content_경변1차 + sam_2_2_content_경변1차 + sam_2_3_content_경변1차) / 3,
                                   int(point2))
        sam_2_stdev_경변1차 = round(((((sam_2_1_content_경변1차 - (
            (sam_2_1_content_경변1차 + sam_2_2_content_경변1차 + sam_2_3_content_경변1차) / 3)) ** 2 + (sam_2_2_content_경변1차 - (
            (sam_2_1_content_경변1차 + sam_2_2_content_경변1차 + sam_2_3_content_경변1차) / 3)) ** 2 + (sam_2_3_content_경변1차 - (
            (sam_2_1_content_경변1차 + sam_2_2_content_경변1차 + sam_2_3_content_경변1차) / 3)) ** 2)) / 2) ** 0.5, 5)

        시료2경변분해율1차 = round(
            ((float(sam_2_average_경변시작) - float(sam_2_average_경변1차)) / float(sam_2_average_경변시작)) * 100, 2)

    if 시료명3_경변1차 == "":
        factor_std3_경변1차 = ""
        sam_3_1_content_경변1차 = ""
        sam_3_2_content_경변1차 = ""
        sam_3_3_content_경변1차 = ""
        sam_3_average_경변1차 = ""
        sam_3_stdev_경변1차 = ""
        시료3경변분해율1차 = ""

    else:
        std_content3_경변1차 = r2[54]
        std_g3_경변1차 = r2[55]
        std_AI_area3_경변1차 = r2[56]
        std_IS_area3_경변1차 = r2[57]

        sam_3_1_g_경변1차 = r2[60]
        sam_3_1_AI_경변1차 = r2[61]
        sam_3_1_IS_경변1차 = r2[62]

        sam_3_2_g_경변1차 = r2[64]
        sam_3_2_AI_경변1차 = r2[65]
        sam_3_2_IS_경변1차 = r2[66]

        sam_3_3_g_경변1차 = r2[68]
        sam_3_3_AI_경변1차 = r2[69]
        sam_3_3_IS_경변1차 = r2[70]

        factor_std3_경변1차 = round(
            float(std_g3_경변1차) * float(std_content3_경변1차) * float(std_IS_area3_경변1차) / float(std_AI_area3_경변1차), 4)
        sam_3_1_content_경변1차 = round(
            (factor_std3_경변1차 * float(sam_3_1_AI_경변1차)) / (float(sam_3_1_IS_경변1차) * (float(sam_3_1_g_경변1차))),
            int(point3))
        sam_3_2_content_경변1차 = round(
            (factor_std3_경변1차 * float(sam_3_2_AI_경변1차)) / (float(sam_3_2_IS_경변1차) * (float(sam_3_2_g_경변1차))),
            int(point3))
        sam_3_3_content_경변1차 = round(
            (factor_std3_경변1차 * float(sam_3_3_AI_경변1차)) / (float(sam_3_3_IS_경변1차) * (float(sam_3_3_g_경변1차))),
            int(point3))
        sam_3_average_경변1차 = round((sam_3_1_content_경변1차 + sam_3_2_content_경변1차 + sam_3_3_content_경변1차) / 3,
                                   int(point3))
        sam_3_stdev_경변1차 = round(((((sam_3_1_content_경변1차 - (
            (sam_3_1_content_경변1차 + sam_3_2_content_경변1차 + sam_3_3_content_경변1차) / 3)) ** 2 + (sam_3_2_content_경변1차 - (
            (sam_3_1_content_경변1차 + sam_3_2_content_경변1차 + sam_3_3_content_경변1차) / 3)) ** 2 + (sam_3_3_content_경변1차 - (
            (sam_3_1_content_경변1차 + sam_3_2_content_경변1차 + sam_3_3_content_경변1차) / 3)) ** 2)) / 2) ** 0.5, 5)

        시료3경변분해율1차 = round(
            ((float(sam_3_average_경변시작) - float(sam_3_average_경변1차)) / float(sam_3_average_경변시작)) * 100, 2)

if r3 == None:
    분석일_경변2차 = ""
    분석년월일_경변2차 = ""
    시료명1_경변2차 = ""
    시료명2_경변2차 = ""
    시료명3_경변2차 = ""
    factor_std1_경변2차 = ""
    sam_1_1_content_경변2차 = ""
    sam_1_2_content_경변2차 = ""
    sam_1_3_content_경변2차 = ""
    sam_1_average_경변2차 = ""
    sam_1_stdev_경변2차 = ""
    시료1경변분해율2차 = ""
    factor_std2_경변2차 = ""
    sam_2_1_content_경변2차 = ""
    sam_2_2_content_경변2차 = ""
    sam_2_3_content_경변2차 = ""
    sam_2_average_경변2차 = ""
    sam_2_stdev_경변2차 = ""
    시료2경변분해율2차 = ""
    factor_std3_경변2차 = ""
    sam_3_1_content_경변2차 = ""
    sam_3_2_content_경변2차 = ""
    sam_3_3_content_경변2차 = ""
    sam_3_average_경변2차 = ""
    sam_3_stdev_경변2차 = ""
    시료3경변분해율2차 = ""

else:
    시료명1_경변2차 = r3[18]
    시료명2_경변2차 = r3[38]
    시료명3_경변2차 = r3[58]
    분석일_경변2차 = r3[9]
    분석년월일_경변2차 = list(re.findall(r"(\d+)", 분석일_경변2차))[0] + "년" + list(re.findall(r"(\d+)", 분석일_경변2차))[
        1] + "월" + list(re.findall(r"(\d+)", 분석일_경변2차))[2] + "일"

    try:
        if 시료명1_경변2차 == None:
            시료명1IUPAC_경변2차 = ""
            시료명1_경변2차 = ""
        else:
            시료명1CID_경변2차 = pcp.get_compounds(시료명1_경변2차, 'name')
            시료명1IUPAC_경변2차 = 시료명1CID_경변2차[0].iupac_name
    except IndexError:
        시료명1IUPAC_경변2차 = ""


    try:
        if 시료명2_경변2차 == None:
            시료명2IUPAC_경변2차 = ""
            시료명2_경변2차 = ""
        else:
            시료명2CID_경변2차 = pcp.get_compounds(시료명2_경변2차, 'name')
            시료명2IUPAC_경변2차 = 시료명2CID_경변2차[0].iupac_name
    except IndexError:
        시료명2IUPAC_경변2차 = ""


    try:
        if 시료명3_경변2차 == None:
            시료명3IUPAC_경변2차 = ""
            시료명3_경변2차 = ""
        else:
            시료명3CID_경변2차 = pcp.get_compounds(시료명3_경변2차, 'name')
            시료명3IUPAC_경변2차 = 시료명3CID_경변2차[0].iupac_name
    except IndexError:
        시료명3IUPAC_경변2차 = ""


    std_content1_경변2차 = r3[14]
    std_g1_경변2차 = r3[15]
    std_AI_area1_경변2차 = r3[16]
    std_IS_area1_경변2차 = r3[17]

    sam_1_1_g_경변2차 = r3[20]
    sam_1_1_AI_경변2차 = r3[21]
    sam_1_1_IS_경변2차 = r3[22]

    sam_1_2_g_경변2차 = r3[24]
    sam_1_2_AI_경변2차 = r3[25]
    sam_1_2_IS_경변2차 = r3[26]

    sam_1_3_g_경변2차 = r3[28]
    sam_1_3_AI_경변2차 = r3[29]
    sam_1_3_IS_경변2차 = r3[30]

    factor_std1_경변2차 = round(
        float(std_g1_경변2차) * float(std_content1_경변2차) * float(std_IS_area1_경변2차) / float(std_AI_area1_경변2차), 4)
    sam_1_1_content_경변2차 = round(
        (factor_std1_경변2차 * float(sam_1_1_AI_경변2차)) / (float(sam_1_1_IS_경변2차) * (float(sam_1_1_g_경변2차))),
        int(point1))
    sam_1_2_content_경변2차 = round(
        (factor_std1_경변2차 * float(sam_1_2_AI_경변2차)) / (float(sam_1_2_IS_경변2차) * (float(sam_1_2_g_경변2차))),
        int(point1))
    sam_1_3_content_경변2차 = round(
        (factor_std1_경변2차 * float(sam_1_3_AI_경변2차)) / (float(sam_1_3_IS_경변2차) * (float(sam_1_3_g_경변2차))),
        int(point1))
    sam_1_average_경변2차 = round((sam_1_1_content_경변2차 + sam_1_2_content_경변2차 + sam_1_3_content_경변2차) / 3,
                               int(point1))
    sam_1_stdev_경변2차 = round(((((sam_1_1_content_경변2차 - (
        (sam_1_1_content_경변2차 + sam_1_2_content_경변2차 + sam_1_3_content_경변2차) / 3)) ** 2 + (
                                    sam_1_2_content_경변2차 - ((
                                                                sam_1_1_content_경변2차 + sam_1_2_content_경변2차 + sam_1_3_content_경변2차) / 3)) ** 2 + (
                                    sam_1_3_content_경변2차 - ((
                                                                sam_1_1_content_경변2차 + sam_1_2_content_경변2차 + sam_1_3_content_경변2차) / 3)) ** 2)) / 2) ** 0.5,
                             5)
    시료1경변분해율2차 = round(
        ((float(sam_1_average_경변시작) - float(sam_1_average_경변2차)) / float(sam_1_average_경변시작)) * 100, 2)

    if 시료명2_경변2차 == "":
        factor_std2_경변2차 = ""
        sam_2_1_content_경변2차 = ""
        sam_2_2_content_경변2차 = ""
        sam_2_3_content_경변2차 = ""
        sam_2_average_경변2차 = ""
        sam_2_stdev_경변2차 = ""
        시료2경변분해율2차 = ""

    else:
        std_content2_경변2차 = r3[34]
        std_g2_경변2차 = r3[35]
        std_AI_area2_경변2차 = r3[36]
        std_IS_area2_경변2차 = r3[37]

        sam_2_1_g_경변2차 = r3[40]
        sam_2_1_AI_경변2차 = r3[41]
        sam_2_1_IS_경변2차 = r3[42]

        sam_2_2_g_경변2차 = r3[44]
        sam_2_2_AI_경변2차 = r3[45]
        sam_2_2_IS_경변2차 = r3[46]

        sam_2_3_g_경변2차 = r3[48]
        sam_2_3_AI_경변2차 = r3[49]
        sam_2_3_IS_경변2차 = r3[50]

        factor_std2_경변2차 = round(
            float(std_g2_경변2차) * float(std_content2_경변2차) * float(std_IS_area2_경변2차) / float(std_AI_area2_경변2차),
            4)
        sam_2_1_content_경변2차 = round(
            (factor_std2_경변2차 * float(sam_2_1_AI_경변2차)) / (float(sam_2_1_IS_경변2차) * (float(sam_2_1_g_경변2차))),
            int(point2))
        sam_2_2_content_경변2차 = round(
            (factor_std2_경변2차 * float(sam_2_2_AI_경변2차)) / (float(sam_2_2_IS_경변2차) * (float(sam_2_2_g_경변2차))),
            int(point2))
        sam_2_3_content_경변2차 = round(
            (factor_std2_경변2차 * float(sam_2_3_AI_경변2차)) / (float(sam_2_3_IS_경변2차) * (float(sam_2_3_g_경변2차))),
            int(point2))
        sam_2_average_경변2차 = round((sam_2_1_content_경변2차 + sam_2_2_content_경변2차 + sam_2_3_content_경변2차) / 3,
                                   int(point2))
        sam_2_stdev_경변2차 = round(((((sam_2_1_content_경변2차 - (
            (sam_2_1_content_경변2차 + sam_2_2_content_경변2차 + sam_2_3_content_경변2차) / 3)) ** 2 + (
                                        sam_2_2_content_경변2차 - (
                                            (
                                                sam_2_1_content_경변2차 + sam_2_2_content_경변2차 + sam_2_3_content_경변2차) / 3)) ** 2 + (
                                        sam_2_3_content_경변2차 - (
                                            (
                                                sam_2_1_content_경변2차 + sam_2_2_content_경변2차 + sam_2_3_content_경변2차) / 3)) ** 2)) / 2) ** 0.5,
                                 5)

        시료2경변분해율2차 = round(
            ((float(sam_2_average_경변시작) - float(sam_2_average_경변2차)) / float(sam_2_average_경변시작)) * 100, 2)

    if 시료명3_경변2차 == "":
        factor_std3_경변2차 = ""
        sam_3_1_content_경변2차 = ""
        sam_3_2_content_경변2차 = ""
        sam_3_3_content_경변2차 = ""
        sam_3_average_경변2차 = ""
        sam_3_stdev_경변2차 = ""
        시료3경변분해율2차 = ""

    else:
        std_content3_경변2차 = r3[54]
        std_g3_경변2차 = r3[55]
        std_AI_area3_경변2차 = r3[56]
        std_IS_area3_경변2차 = r3[57]

        sam_3_1_g_경변2차 = r3[60]
        sam_3_1_AI_경변2차 = r3[61]
        sam_3_1_IS_경변2차 = r3[62]

        sam_3_2_g_경변2차 = r3[64]
        sam_3_2_AI_경변2차 = r3[65]
        sam_3_2_IS_경변2차 = r3[66]

        sam_3_3_g_경변2차 = r3[68]
        sam_3_3_AI_경변2차 = r3[69]
        sam_3_3_IS_경변2차 = r3[70]

        factor_std3_경변2차 = round(
            float(std_g3_경변2차) * float(std_content3_경변2차) * float(std_IS_area3_경변2차) / float(std_AI_area3_경변2차),
            4)
        sam_3_1_content_경변2차 = round(
            (factor_std3_경변2차 * float(sam_3_1_AI_경변2차)) / (float(sam_3_1_IS_경변2차) * (float(sam_3_1_g_경변2차))),
            int(point3))
        sam_3_2_content_경변2차 = round(
            (factor_std3_경변2차 * float(sam_3_2_AI_경변2차)) / (float(sam_3_2_IS_경변2차) * (float(sam_3_2_g_경변2차))),
            int(point3))
        sam_3_3_content_경변2차 = round(
            (factor_std3_경변2차 * float(sam_3_3_AI_경변2차)) / (float(sam_3_3_IS_경변2차) * (float(sam_3_3_g_경변2차))),
            int(point3))
        sam_3_average_경변2차 = round((sam_3_1_content_경변2차 + sam_3_2_content_경변2차 + sam_3_3_content_경변2차) / 3,
                                   int(point3))
        sam_3_stdev_경변2차 = round(((((sam_3_1_content_경변2차 - (
            (sam_3_1_content_경변2차 + sam_3_2_content_경변2차 + sam_3_3_content_경변2차) / 3)) ** 2 + (
                                        sam_3_2_content_경변2차 - (
                                            (
                                                sam_3_1_content_경변2차 + sam_3_2_content_경변2차 + sam_3_3_content_경변2차) / 3)) ** 2 + (
                                        sam_3_3_content_경변2차 - (
                                            (
                                                sam_3_1_content_경변2차 + sam_3_2_content_경변2차 + sam_3_3_content_경변2차) / 3)) ** 2)) / 2) ** 0.5,
                                 5)

        시료3경변분해율2차 = round(
            ((float(sam_3_average_경변시작) - float(sam_3_average_경변2차)) / float(sam_3_average_경변시작)) * 100, 2)

if r4 == None:
    분석일_경변3차 = ""
    분석년월일_경변3차 = ""
    시료명1_경변3차 = ""
    시료명2_경변3차 = ""
    시료명3_경변3차 = ""
    factor_std1_경변3차 = ""
    sam_1_1_content_경변3차 = ""
    sam_1_2_content_경변3차 = ""
    sam_1_3_content_경변3차 = ""
    sam_1_average_경변3차 = ""
    sam_1_stdev_경변3차 = ""
    시료1경변분해율3차 = ""
    factor_std2_경변3차 = ""
    sam_2_1_content_경변3차 = ""
    sam_2_2_content_경변3차 = ""
    sam_2_3_content_경변3차 = ""
    sam_2_average_경변3차 = ""
    sam_2_stdev_경변3차 = ""
    시료2경변분해율3차 = ""
    factor_std3_경변3차 = ""
    sam_3_1_content_경변3차 = ""
    sam_3_2_content_경변3차 = ""
    sam_3_3_content_경변3차 = ""
    sam_3_average_경변3차 = ""
    sam_3_stdev_경변3차 = ""
    시료3경변분해율3차 = ""

else:
    시료명1_경변3차 = r4[18]
    시료명2_경변3차 = r4[38]
    시료명3_경변3차 = r4[58]
    분석일_경변3차 = r4[9]
    분석년월일_경변3차 = list(re.findall(r"(\d+)", 분석일_경변3차))[0] + "년" + list(re.findall(r"(\d+)", 분석일_경변3차))[
        1] + "월" + list(re.findall(r"(\d+)", 분석일_경변3차))[2] + "일"

    try:
        if 시료명1_경변3차 == None:
            시료명1IUPAC_경변3차 = ""
            시료명1_경변3차 = ""
        else:
            시료명1CID_경변3차 = pcp.get_compounds(시료명1_경변3차, 'name')
            시료명1IUPAC_경변3차 = 시료명1CID_경변3차[0].iupac_name
    except IndexError:
        시료명1IUPAC_경변3차 = ""


    try:
        if 시료명2_경변3차 == None:
            시료명2IUPAC_경변3차 = ""
            시료명2_경변3차 = ""
        else:
            시료명2CID_경변3차 = pcp.get_compounds(시료명2_경변3차, 'name')
            시료명2IUPAC_경변3차 = 시료명2CID_경변3차[0].iupac_name
    except IndexError:
        시료명2IUPAC_경변3차 = ""


    try:
        if 시료명3_경변3차 == None:
            시료명3IUPAC_경변3차 = ""
            시료명3_경변3차 = ""
        else:
            시료명3CID_경변3차 = pcp.get_compounds(시료명3_경변3차, 'name')
            시료명3IUPAC_경변3차 = 시료명3CID_경변3차[0].iupac_name
    except IndexError:
        시료명3IUPAC_경변3차 = ""


    std_content1_경변3차 = r4[14]
    std_g1_경변3차 = r4[15]
    std_AI_area1_경변3차 = r4[16]
    std_IS_area1_경변3차 = r4[17]

    sam_1_1_g_경변3차 = r4[20]
    sam_1_1_AI_경변3차 = r4[21]
    sam_1_1_IS_경변3차 = r4[22]

    sam_1_2_g_경변3차 = r4[24]
    sam_1_2_AI_경변3차 = r4[25]
    sam_1_2_IS_경변3차 = r4[26]

    sam_1_3_g_경변3차 = r4[28]
    sam_1_3_AI_경변3차 = r4[29]
    sam_1_3_IS_경변3차 = r4[30]

    factor_std1_경변3차 = round(
        float(std_g1_경변3차) * float(std_content1_경변3차) * float(std_IS_area1_경변3차) / float(std_AI_area1_경변3차), 4)
    sam_1_1_content_경변3차 = round(
        (factor_std1_경변3차 * float(sam_1_1_AI_경변3차)) / (float(sam_1_1_IS_경변3차) * (float(sam_1_1_g_경변3차))),
        int(point1))
    sam_1_2_content_경변3차 = round(
        (factor_std1_경변3차 * float(sam_1_2_AI_경변3차)) / (float(sam_1_2_IS_경변3차) * (float(sam_1_2_g_경변3차))),
        int(point1))
    sam_1_3_content_경변3차 = round(
        (factor_std1_경변3차 * float(sam_1_3_AI_경변3차)) / (float(sam_1_3_IS_경변3차) * (float(sam_1_3_g_경변3차))),
        int(point1))
    sam_1_average_경변3차 = round((sam_1_1_content_경변3차 + sam_1_2_content_경변3차 + sam_1_3_content_경변3차) / 3,
                               int(point1))
    sam_1_stdev_경변3차 = round(((((sam_1_1_content_경변3차 - (
        (sam_1_1_content_경변3차 + sam_1_2_content_경변3차 + sam_1_3_content_경변3차) / 3)) ** 2 + (
                                    sam_1_2_content_경변3차 - ((
                                                                sam_1_1_content_경변3차 + sam_1_2_content_경변3차 + sam_1_3_content_경변3차) / 3)) ** 2 + (
                                    sam_1_3_content_경변3차 - ((
                                                                sam_1_1_content_경변3차 + sam_1_2_content_경변3차 + sam_1_3_content_경변3차) / 3)) ** 2)) / 2) ** 0.5,
                             5)
    시료1경변분해율3차 = round(
        ((float(sam_1_average_경변시작) - float(sam_1_average_경변3차)) / float(sam_1_average_경변시작)) * 100, 2)

    if 시료명2_경변3차 == "":
        factor_std2_경변3차 = ""
        sam_2_1_content_경변3차 = ""
        sam_2_2_content_경변3차 = ""
        sam_2_3_content_경변3차 = ""
        sam_2_average_경변3차 = ""
        sam_2_stdev_경변3차 = ""
        시료2경변분해율3차 = ""

    else:
        std_content2_경변3차 = r4[34]
        std_g2_경변3차 = r4[35]
        std_AI_area2_경변3차 = r4[36]
        std_IS_area2_경변3차 = r4[37]

        sam_2_1_g_경변3차 = r4[40]
        sam_2_1_AI_경변3차 = r4[41]
        sam_2_1_IS_경변3차 = r4[42]

        sam_2_2_g_경변3차 = r4[44]
        sam_2_2_AI_경변3차 = r4[45]
        sam_2_2_IS_경변3차 = r4[46]

        sam_2_3_g_경변3차 = r4[48]
        sam_2_3_AI_경변3차 = r4[49]
        sam_2_3_IS_경변3차 = r4[50]

        factor_std2_경변3차 = round(
            float(std_g2_경변3차) * float(std_content2_경변3차) * float(std_IS_area2_경변3차) / float(std_AI_area2_경변3차),
            4)
        sam_2_1_content_경변3차 = round(
            (factor_std2_경변3차 * float(sam_2_1_AI_경변3차)) / (float(sam_2_1_IS_경변3차) * (float(sam_2_1_g_경변3차))),
            int(point2))
        sam_2_2_content_경변3차 = round(
            (factor_std2_경변3차 * float(sam_2_2_AI_경변3차)) / (float(sam_2_2_IS_경변3차) * (float(sam_2_2_g_경변3차))),
            int(point2))
        sam_2_3_content_경변3차 = round(
            (factor_std2_경변3차 * float(sam_2_3_AI_경변3차)) / (float(sam_2_3_IS_경변3차) * (float(sam_2_3_g_경변3차))),
            int(point2))
        sam_2_average_경변3차 = round((sam_2_1_content_경변3차 + sam_2_2_content_경변3차 + sam_2_3_content_경변3차) / 3,
                                   int(point2))
        sam_2_stdev_경변3차 = round(((((sam_2_1_content_경변3차 - (
            (sam_2_1_content_경변3차 + sam_2_2_content_경변3차 + sam_2_3_content_경변3차) / 3)) ** 2 + (
                                        sam_2_2_content_경변3차 - (
                                            (
                                                sam_2_1_content_경변3차 + sam_2_2_content_경변3차 + sam_2_3_content_경변3차) / 3)) ** 2 + (
                                        sam_2_3_content_경변3차 - (
                                            (
                                                sam_2_1_content_경변3차 + sam_2_2_content_경변3차 + sam_2_3_content_경변3차) / 3)) ** 2)) / 2) ** 0.5,
                                 5)

        시료2경변분해율3차 = round(
            ((float(sam_2_average_경변시작) - float(sam_2_average_경변3차)) / float(sam_2_average_경변시작)) * 100, 2)

    if 시료명3_경변3차 == "":
        factor_std3_경변3차 = ""
        sam_3_1_content_경변3차 = ""
        sam_3_2_content_경변3차 = ""
        sam_3_3_content_경변3차 = ""
        sam_3_average_경변3차 = ""
        sam_3_stdev_경변3차 = ""
        시료3경변분해율3차 = ""

    else:
        std_content3_경변3차 = r4[54]
        std_g3_경변3차 = r4[55]
        std_AI_area3_경변3차 = r4[56]
        std_IS_area3_경변3차 = r4[57]

        sam_3_1_g_경변3차 = r4[60]
        sam_3_1_AI_경변3차 = r4[61]
        sam_3_1_IS_경변3차 = r4[62]

        sam_3_2_g_경변3차 = r4[64]
        sam_3_2_AI_경변3차 = r4[65]
        sam_3_2_IS_경변3차 = r4[66]

        sam_3_3_g_경변3차 = r4[68]
        sam_3_3_AI_경변3차 = r4[69]
        sam_3_3_IS_경변3차 = r4[70]

        factor_std3_경변3차 = round(
            float(std_g3_경변3차) * float(std_content3_경변3차) * float(std_IS_area3_경변3차) / float(std_AI_area3_경변3차),
            4)
        sam_3_1_content_경변3차 = round(
            (factor_std3_경변3차 * float(sam_3_1_AI_경변3차)) / (float(sam_3_1_IS_경변3차) * (float(sam_3_1_g_경변3차))),
            int(point3))
        sam_3_2_content_경변3차 = round(
            (factor_std3_경변3차 * float(sam_3_2_AI_경변3차)) / (float(sam_3_2_IS_경변3차) * (float(sam_3_2_g_경변3차))),
            int(point3))
        sam_3_3_content_경변3차 = round(
            (factor_std3_경변3차 * float(sam_3_3_AI_경변3차)) / (float(sam_3_3_IS_경변3차) * (float(sam_3_3_g_경변3차))),
            int(point3))
        sam_3_average_경변3차 = round((sam_3_1_content_경변3차 + sam_3_2_content_경변3차 + sam_3_3_content_경변3차) / 3,
                                   int(point3))
        sam_3_stdev_경변3차 = round(((((sam_3_1_content_경변3차 - (
            (sam_3_1_content_경변3차 + sam_3_2_content_경변3차 + sam_3_3_content_경변3차) / 3)) ** 2 + (
                                        sam_3_2_content_경변3차 - (
                                            (
                                                sam_3_1_content_경변3차 + sam_3_2_content_경변3차 + sam_3_3_content_경변3차) / 3)) ** 2 + (
                                        sam_3_3_content_경변3차 - (
                                            (
                                                sam_3_1_content_경변3차 + sam_3_2_content_경변3차 + sam_3_3_content_경변3차) / 3)) ** 2)) / 2) ** 0.5,
                                 5)

        시료3경변분해율3차 = round(
            ((float(sam_3_average_경변시작) - float(sam_3_average_경변3차)) / float(sam_3_average_경변시작)) * 100, 2)

if r5 == None:
    분석일_경변4차 = ""
    분석년월일_경변4차 = ""
    시료명1_경변4차 = ""
    시료명2_경변4차 = ""
    시료명3_경변4차 = ""
    factor_std1_경변4차 = ""
    sam_1_1_content_경변4차 = ""
    sam_1_2_content_경변4차 = ""
    sam_1_3_content_경변4차 = ""
    sam_1_average_경변4차 = ""
    sam_1_stdev_경변4차 = ""
    시료1경변분해율4차 = ""
    factor_std2_경변4차 = ""
    sam_2_1_content_경변4차 = ""
    sam_2_2_content_경변4차 = ""
    sam_2_3_content_경변4차 = ""
    sam_2_average_경변4차 = ""
    sam_2_stdev_경변4차 = ""
    시료2경변분해율4차 = ""
    factor_std3_경변4차 = ""
    sam_3_1_content_경변4차 = ""
    sam_3_2_content_경변4차 = ""
    sam_3_3_content_경변4차 = ""
    sam_3_average_경변4차 = ""
    sam_3_stdev_경변4차 = ""
    시료3경변분해율4차 = ""

else:
    시료명1_경변4차 = r5[18]
    시료명2_경변4차 = r5[38]
    시료명3_경변4차 = r5[58]
    분석일_경변4차 = r5[9]
    분석년월일_경변4차 = list(re.findall(r"(\d+)", 분석일_경변4차))[0] + "년" + list(re.findall(r"(\d+)", 분석일_경변4차))[
        1] + "월" + list(re.findall(r"(\d+)", 분석일_경변4차))[2] + "일"

    try:
        if 시료명1_경변4차 == None:
            시료명1IUPAC_경변4차 = ""
            시료명1_경변4차 = ""
        else:
            시료명1CID_경변4차 = pcp.get_compounds(시료명1_경변4차, 'name')
            시료명1IUPAC_경변4차 = 시료명1CID_경변4차[0].iupac_name
    except IndexError:
        시료명1IUPAC_경변4차 = ""


    try:
        if 시료명2_경변4차 == None:
            시료명2IUPAC_경변4차 = ""
            시료명2_경변4차 = ""
        else:
            시료명2CID_경변4차 = pcp.get_compounds(시료명2_경변4차, 'name')
            시료명2IUPAC_경변4차 = 시료명2CID_경변4차[0].iupac_name
    except IndexError:
        시료명2IUPAC_경변4차 = ""


    try:
        if 시료명3_경변4차 == None:
            시료명3IUPAC_경변4차 = ""
            시료명3_경변4차 = ""
        else:
            시료명3CID_경변4차 = pcp.get_compounds(시료명3_경변4차, 'name')
            시료명3IUPAC_경변4차 = 시료명3CID_경변4차[0].iupac_name
    except IndexError:
        시료명3IUPAC_경변4차 = ""


    std_content1_경변4차 = r5[14]
    std_g1_경변4차 = r5[15]
    std_AI_area1_경변4차 = r5[16]
    std_IS_area1_경변4차 = r5[17]

    sam_1_1_g_경변4차 = r5[20]
    sam_1_1_AI_경변4차 = r5[21]
    sam_1_1_IS_경변4차 = r5[22]

    sam_1_2_g_경변4차 = r5[24]
    sam_1_2_AI_경변4차 = r5[25]
    sam_1_2_IS_경변4차 = r5[26]

    sam_1_3_g_경변4차 = r5[28]
    sam_1_3_AI_경변4차 = r5[29]
    sam_1_3_IS_경변4차 = r5[30]

    factor_std1_경변4차 = round(
        float(std_g1_경변4차) * float(std_content1_경변4차) * float(std_IS_area1_경변4차) / float(std_AI_area1_경변4차), 4)
    sam_1_1_content_경변4차 = round(
        (factor_std1_경변4차 * float(sam_1_1_AI_경변4차)) / (float(sam_1_1_IS_경변4차) * (float(sam_1_1_g_경변4차))),
        int(point1))
    sam_1_2_content_경변4차 = round(
        (factor_std1_경변4차 * float(sam_1_2_AI_경변4차)) / (float(sam_1_2_IS_경변4차) * (float(sam_1_2_g_경변4차))),
        int(point1))
    sam_1_3_content_경변4차 = round(
        (factor_std1_경변4차 * float(sam_1_3_AI_경변4차)) / (float(sam_1_3_IS_경변4차) * (float(sam_1_3_g_경변4차))),
        int(point1))
    sam_1_average_경변4차 = round((sam_1_1_content_경변4차 + sam_1_2_content_경변4차 + sam_1_3_content_경변4차) / 3,
                               int(point1))
    sam_1_stdev_경변4차 = round(((((sam_1_1_content_경변4차 - (
        (sam_1_1_content_경변4차 + sam_1_2_content_경변4차 + sam_1_3_content_경변4차) / 3)) ** 2 + (
                                    sam_1_2_content_경변4차 - ((
                                                                sam_1_1_content_경변4차 + sam_1_2_content_경변4차 + sam_1_3_content_경변4차) / 3)) ** 2 + (
                                    sam_1_3_content_경변4차 - ((
                                                                sam_1_1_content_경변4차 + sam_1_2_content_경변4차 + sam_1_3_content_경변4차) / 3)) ** 2)) / 2) ** 0.5,
                             5)
    시료1경변분해율4차 = round(
        ((float(sam_1_average_경변시작) - float(sam_1_average_경변4차)) / float(sam_1_average_경변시작)) * 100, 2)

    if 시료명2_경변4차 == "":
        factor_std2_경변4차 = ""
        sam_2_1_content_경변4차 = ""
        sam_2_2_content_경변4차 = ""
        sam_2_3_content_경변4차 = ""
        sam_2_average_경변4차 = ""
        sam_2_stdev_경변4차 = ""
        시료2경변분해율4차 = ""

    else:
        std_content2_경변4차 = r5[34]
        std_g2_경변4차 = r5[35]
        std_AI_area2_경변4차 = r5[36]
        std_IS_area2_경변4차 = r5[37]

        sam_2_1_g_경변4차 = r5[40]
        sam_2_1_AI_경변4차 = r5[41]
        sam_2_1_IS_경변4차 = r5[42]

        sam_2_2_g_경변4차 = r5[44]
        sam_2_2_AI_경변4차 = r5[45]
        sam_2_2_IS_경변4차 = r5[46]

        sam_2_3_g_경변4차 = r5[48]
        sam_2_3_AI_경변4차 = r5[49]
        sam_2_3_IS_경변4차 = r5[50]

        factor_std2_경변4차 = round(
            float(std_g2_경변4차) * float(std_content2_경변4차) * float(std_IS_area2_경변4차) / float(std_AI_area2_경변4차), 4)
        sam_2_1_content_경변4차 = round(
            (factor_std2_경변4차 * float(sam_2_1_AI_경변4차)) / (float(sam_2_1_IS_경변4차) * (float(sam_2_1_g_경변4차))),
            int(point2))
        sam_2_2_content_경변4차 = round(
            (factor_std2_경변4차 * float(sam_2_2_AI_경변4차)) / (float(sam_2_2_IS_경변4차) * (float(sam_2_2_g_경변4차))),
            int(point2))
        sam_2_3_content_경변4차 = round(
            (factor_std2_경변4차 * float(sam_2_3_AI_경변4차)) / (float(sam_2_3_IS_경변4차) * (float(sam_2_3_g_경변4차))),
            int(point2))
        sam_2_average_경변4차 = round((sam_2_1_content_경변4차 + sam_2_2_content_경변4차 + sam_2_3_content_경변4차) / 3,
                                   int(point2))
        sam_2_stdev_경변4차 = round(((((sam_2_1_content_경변4차 - (
            (sam_2_1_content_경변4차 + sam_2_2_content_경변4차 + sam_2_3_content_경변4차) / 3)) ** 2 + (
                                        sam_2_2_content_경변4차 - (
                                            (
                                            sam_2_1_content_경변4차 + sam_2_2_content_경변4차 + sam_2_3_content_경변4차) / 3)) ** 2 + (
                                        sam_2_3_content_경변4차 - ((
                                                                    sam_2_1_content_경변4차 + sam_2_2_content_경변4차 + sam_2_3_content_경변4차) / 3)) ** 2)) / 2) ** 0.5,
                                 5)

        시료2경변분해율4차 = round(
            ((float(sam_2_average_경변시작) - float(sam_2_average_경변4차)) / float(sam_2_average_경변시작)) * 100, 2)

    if 시료명3_경변4차 == "":
        factor_std3_경변4차 = ""
        sam_3_1_content_경변4차 = ""
        sam_3_2_content_경변4차 = ""
        sam_3_3_content_경변4차 = ""
        sam_3_average_경변4차 = ""
        sam_3_stdev_경변4차 = ""
        시료3경변분해율4차 = ""

    else:
        std_content3_경변4차 = r5[54]
        std_g3_경변4차 = r5[55]
        std_AI_area3_경변4차 = r5[56]
        std_IS_area3_경변4차 = r5[57]

        sam_3_1_g_경변4차 = r5[60]
        sam_3_1_AI_경변4차 = r5[61]
        sam_3_1_IS_경변4차 = r5[62]

        sam_3_2_g_경변4차 = r5[64]
        sam_3_2_AI_경변4차 = r5[65]
        sam_3_2_IS_경변4차 = r5[66]

        sam_3_3_g_경변4차 = r5[68]
        sam_3_3_AI_경변4차 = r5[69]
        sam_3_3_IS_경변4차 = r5[70]

        factor_std3_경변4차 = round(
            float(std_g3_경변4차) * float(std_content3_경변4차) * float(std_IS_area3_경변4차) / float(std_AI_area3_경변4차),
            4)
        sam_3_1_content_경변4차 = round(
            (factor_std3_경변4차 * float(sam_3_1_AI_경변4차)) / (float(sam_3_1_IS_경변4차) * (float(sam_3_1_g_경변4차))),
            int(point3))
        sam_3_2_content_경변4차 = round(
            (factor_std3_경변4차 * float(sam_3_2_AI_경변4차)) / (float(sam_3_2_IS_경변4차) * (float(sam_3_2_g_경변4차))),
            int(point3))
        sam_3_3_content_경변4차 = round(
            (factor_std3_경변4차 * float(sam_3_3_AI_경변4차)) / (float(sam_3_3_IS_경변4차) * (float(sam_3_3_g_경변4차))),
            int(point3))
        sam_3_average_경변4차 = round((sam_3_1_content_경변4차 + sam_3_2_content_경변4차 + sam_3_3_content_경변4차) / 3,
                                   int(point3))
        sam_3_stdev_경변4차 = round(((((sam_3_1_content_경변4차 - (
            (sam_3_1_content_경변4차 + sam_3_2_content_경변4차 + sam_3_3_content_경변4차) / 3)) ** 2 + (
                                        sam_3_2_content_경변4차 - (
                                            (
                                            sam_3_1_content_경변4차 + sam_3_2_content_경변4차 + sam_3_3_content_경변4차) / 3)) ** 2 + (
                                        sam_3_3_content_경변4차 - (
                                            (
                                            sam_3_1_content_경변4차 + sam_3_2_content_경변4차 + sam_3_3_content_경변4차) / 3)) ** 2)) / 2) ** 0.5,
                                 5)

        시료3경변분해율4차 = round(
            ((float(sam_3_average_경변시작) - float(sam_3_average_경변4차)) / float(sam_3_average_경변시작)) * 100, 2)

if r6 == None:
    분석일_저온 = ""
    분석년월일_저온 = ""
    시료명1_저온 = ""
    시료명2_저온 = ""
    시료명3_저온 = ""
    factor_std1_저온 = ""
    sam_1_1_content_저온 = ""
    sam_1_2_content_저온 = ""
    sam_1_3_content_저온 = ""
    sam_1_average_저온 = ""
    sam_1_stdev_저온 = ""
    시료1경변분해율저온 = ""
    factor_std2_저온 = ""
    sam_2_1_content_저온 = ""
    sam_2_2_content_저온 = ""
    sam_2_3_content_저온 = ""
    sam_2_average_저온 = ""
    sam_2_stdev_저온 = ""
    시료2경변분해율저온 = ""
    factor_std3_저온 = ""
    sam_3_1_content_저온 = ""
    sam_3_2_content_저온 = ""
    sam_3_3_content_저온 = ""
    sam_3_average_저온 = ""
    sam_3_stdev_저온 = ""
    시료3경변분해율저온 = ""

else:
    시료명1_저온 = r6[18]
    시료명2_저온 = r6[38]
    시료명3_저온 = r6[58]
    분석일_저온 = r6[9]
    분석년월일_저온 = list(re.findall(r"(\d+)", 분석일_저온))[0] + "년" + list(re.findall(r"(\d+)", 분석일_저온))[
        1] + "월" + list(re.findall(r"(\d+)", 분석일_저온))[2] + "일"

    try:
        if 시료명1_저온 == None:
            시료명1IUPAC_저온 = ""
            시료명1_저온 = ""
        else:
            시료명1CID_저온 = pcp.get_compounds(시료명1_저온, 'name')
            시료명1IUPAC_저온 = 시료명1CID_저온[0].iupac_name
    except IndexError:
        시료명1IUPAC_저온 = ""


    try:
        if 시료명2_저온 == None:
            시료명2IUPAC_저온 = ""
            시료명2_저온 = ""
        else:
            시료명2CID_저온 = pcp.get_compounds(시료명2_저온, 'name')
            시료명2IUPAC_저온 = 시료명2CID_저온[0].iupac_name
    except IndexError:
        시료명2IUPAC_저온 = ""


    try:
        if 시료명3_저온 == None:
            시료명3IUPAC_저온 = ""
            시료명3_저온 = ""
        else:
            시료명3CID_저온 = pcp.get_compounds(시료명3_저온, 'name')
            시료명3IUPAC_저온 = 시료명3CID_저온[0].iupac_name
    except IndexError:
        시료명3IUPAC_저온 = ""


    std_content1_저온 = r6[14]
    std_g1_저온 = r6[15]
    std_AI_area1_저온 = r6[16]
    std_IS_area1_저온 = r6[17]

    sam_1_1_g_저온 = r6[20]
    sam_1_1_AI_저온 = r6[21]
    sam_1_1_IS_저온 = r6[22]

    sam_1_2_g_저온 = r6[24]
    sam_1_2_AI_저온 = r6[25]
    sam_1_2_IS_저온 = r6[26]

    sam_1_3_g_저온 = r6[28]
    sam_1_3_AI_저온 = r6[29]
    sam_1_3_IS_저온 = r6[30]

    factor_std1_저온 = round(
        float(std_g1_저온) * float(std_content1_저온) * float(std_IS_area1_저온) / float(std_AI_area1_저온), 4)
    sam_1_1_content_저온 = round(
        (factor_std1_저온 * float(sam_1_1_AI_저온)) / (float(sam_1_1_IS_저온) * (float(sam_1_1_g_저온))),
        int(point1))
    sam_1_2_content_저온 = round(
        (factor_std1_저온 * float(sam_1_2_AI_저온)) / (float(sam_1_2_IS_저온) * (float(sam_1_2_g_저온))),
        int(point1))
    sam_1_3_content_저온 = round(
        (factor_std1_저온 * float(sam_1_3_AI_저온)) / (float(sam_1_3_IS_저온) * (float(sam_1_3_g_저온))),
        int(point1))
    sam_1_average_저온 = round((sam_1_1_content_저온 + sam_1_2_content_저온 + sam_1_3_content_저온) / 3,
                             int(point1))
    sam_1_stdev_저온 = round(((((sam_1_1_content_저온 - (
        (sam_1_1_content_저온 + sam_1_2_content_저온 + sam_1_3_content_저온) / 3)) ** 2 + (
                                  sam_1_2_content_저온 - ((
                                                            sam_1_1_content_저온 + sam_1_2_content_저온 + sam_1_3_content_저온) / 3)) ** 2 + (
                                  sam_1_3_content_저온 - ((
                                                            sam_1_1_content_저온 + sam_1_2_content_저온 + sam_1_3_content_저온) / 3)) ** 2)) / 2) ** 0.5,
                           5)
    시료1경변분해율저온 = round(
        ((float(sam_1_average_경변시작) - float(sam_1_average_저온)) / float(sam_1_average_경변시작)) * 100, 2)

    if 시료명2_저온 == "":
        factor_std2_저온 = ""
        sam_2_1_content_저온 = ""
        sam_2_2_content_저온 = ""
        sam_2_3_content_저온 = ""
        sam_2_average_저온 = ""
        sam_2_stdev_저온 = ""
        시료2경변분해율저온 = ""

    else:
        std_content2_저온 = r6[34]
        std_g2_저온 = r6[35]
        std_AI_area2_저온 = r6[36]
        std_IS_area2_저온 = r6[37]

        sam_2_1_g_저온 = r6[40]
        sam_2_1_AI_저온 = r6[41]
        sam_2_1_IS_저온 = r6[42]

        sam_2_2_g_저온 = r6[44]
        sam_2_2_AI_저온 = r6[45]
        sam_2_2_IS_저온 = r6[46]

        sam_2_3_g_저온 = r6[48]
        sam_2_3_AI_저온 = r6[49]
        sam_2_3_IS_저온 = r6[50]

        factor_std2_저온 = round(
            float(std_g2_저온) * float(std_content2_저온) * float(std_IS_area2_저온) / float(std_AI_area2_저온),
            4)
        sam_2_1_content_저온 = round(
            (factor_std2_저온 * float(sam_2_1_AI_저온)) / (float(sam_2_1_IS_저온) * (float(sam_2_1_g_저온))),
            int(point2))
        sam_2_2_content_저온 = round(
            (factor_std2_저온 * float(sam_2_2_AI_저온)) / (float(sam_2_2_IS_저온) * (float(sam_2_2_g_저온))),
            int(point2))
        sam_2_3_content_저온 = round(
            (factor_std2_저온 * float(sam_2_3_AI_저온)) / (float(sam_2_3_IS_저온) * (float(sam_2_3_g_저온))),
            int(point2))
        sam_2_average_저온 = round((sam_2_1_content_저온 + sam_2_2_content_저온 + sam_2_3_content_저온) / 3,
                                 int(point2))
        sam_2_stdev_저온 = round(((((sam_2_1_content_저온 - (
            (sam_2_1_content_저온 + sam_2_2_content_저온 + sam_2_3_content_저온) / 3)) ** 2 + (
                                      sam_2_2_content_저온 - (
                                          (
                                              sam_2_1_content_저온 + sam_2_2_content_저온 + sam_2_3_content_저온) / 3)) ** 2 + (
                                      sam_2_3_content_저온 - (
                                          (
                                              sam_2_1_content_저온 + sam_2_2_content_저온 + sam_2_3_content_저온) / 3)) ** 2)) / 2) ** 0.5,
                               5)

        시료2경변분해율저온 = round(
            ((float(sam_2_average_경변시작) - float(sam_2_average_저온)) / float(sam_2_average_경변시작)) * 100, 2)

    if 시료명3_저온 == "":
        factor_std3_저온 = ""
        sam_3_1_content_저온 = ""
        sam_3_2_content_저온 = ""
        sam_3_3_content_저온 = ""
        sam_3_average_저온 = ""
        sam_3_stdev_저온 = ""
        시료3경변분해율저온 = ""

    else:
        std_content3_저온 = r6[54]
        std_g3_저온 = r6[55]
        std_AI_area3_저온 = r6[56]
        std_IS_area3_저온 = r6[57]

        sam_3_1_g_저온 = r6[60]
        sam_3_1_AI_저온 = r6[61]
        sam_3_1_IS_저온 = r6[62]

        sam_3_2_g_저온 = r6[64]
        sam_3_2_AI_저온 = r6[65]
        sam_3_2_IS_저온 = r6[66]

        sam_3_3_g_저온 = r6[68]
        sam_3_3_AI_저온 = r6[69]
        sam_3_3_IS_저온 = r6[70]

        factor_std3_저온 = round(
            float(std_g3_저온) * float(std_content3_저온) * float(std_IS_area3_저온) / float(std_AI_area3_저온),
            4)
        sam_3_1_content_저온 = round(
            (factor_std3_저온 * float(sam_3_1_AI_저온)) / (float(sam_3_1_IS_저온) * (float(sam_3_1_g_저온))),
            int(point3))
        sam_3_2_content_저온 = round(
            (factor_std3_저온 * float(sam_3_2_AI_저온)) / (float(sam_3_2_IS_저온) * (float(sam_3_2_g_저온))),
            int(point3))
        sam_3_3_content_저온 = round(
            (factor_std3_저온 * float(sam_3_3_AI_저온)) / (float(sam_3_3_IS_저온) * (float(sam_3_3_g_저온))),
            int(point3))
        sam_3_average_저온 = round((sam_3_1_content_저온 + sam_3_2_content_저온 + sam_3_3_content_저온) / 3,
                                 int(point3))
        sam_3_stdev_저온 = round(((((sam_3_1_content_저온 - (
            (sam_3_1_content_저온 + sam_3_2_content_저온 + sam_3_3_content_저온) / 3)) ** 2 + (sam_3_2_content_저온 - ((sam_3_1_content_저온 + sam_3_2_content_저온 + sam_3_3_content_저온) / 3)) ** 2 + (
                                      sam_3_3_content_저온 - ((sam_3_1_content_저온 + sam_3_2_content_저온 + sam_3_3_content_저온) / 3)) ** 2)) / 2) ** 0.5, 5)

        시료3경변분해율저온 = round(
            ((float(sam_3_average_경변시작) - float(sam_3_average_저온)) / float(sam_3_average_경변시작)) * 100, 2)

if r3 == None:
    시험기간 = 분석일_저온

elif r4 == None:
    시험기간 = 분석일_경변2차

elif r5 == None:
    시험기간 = 분석일_경변3차

else:
    시험기간 = 분석일_경변4차



print('Lot No.: ', LOTNO_경변시작, ' 제조년월일:', 제조년_경변시작 + '년', 제조월_경변시작 + '월', 제조일_경변시작 + '일')
print('분석책임자 (주)팜한농 작물보호연구센터: ', 책임자)
print('시험의뢰자 (주)팜한농: ', 의뢰자)
print('품목명: ', 한글명1_경변시작, 한글명2_경변시작, 한글명3_경변시작, 제형분류_경변시작)
print('영문명: ', 시료명1_경변시작, 시료명2_경변시작, 시료명3_경변시작)
print('유효성분의 명칭 및 함유량 :')
print(시료명1IUPAC_경변시작, ':', 함량1_경변시작)
print(시료명2IUPAC_경변시작, ':', 함량2_경변시작)
print(시료명3IUPAC_경변시작, ':', 함량3_경변시작)
print('시험기간: ', 분석년월일_경변시작, '~', list(re.findall(r"(\d+)", 시험기간))[0] + "년 ", list(re.findall(r"(\d+)", 시험기간))[1] + "월 ",
      list(re.findall(r"(\d+)", 시험기간))[2] + "일")
print('포장용기 및 재질: ')

print('유효성분 함량(시료1, 시작):', sam_1_1_content_경변시작, sam_1_2_content_경변시작, sam_1_3_content_경변시작, sam_1_average_경변시작, sam_1_stdev_경변시작)
print('유효성분 함량(시료2, 시작):', sam_2_1_content_경변시작, sam_2_2_content_경변시작, sam_2_3_content_경변시작, sam_2_average_경변시작, sam_2_stdev_경변시작)
print('유효성분 함량(시료3, 시작):', sam_3_1_content_경변시작, sam_3_2_content_경변시작, sam_3_3_content_경변시작, sam_3_average_경변시작, sam_3_stdev_경변시작)

print('유효성분 함량(시료1, 1년차):', sam_1_1_content_경변1차, sam_1_2_content_경변1차, sam_1_3_content_경변1차, sam_1_average_경변1차, sam_1_stdev_경변1차)
print('유효성분 함량(시료2, 1년차):', sam_2_1_content_경변1차, sam_2_2_content_경변1차, sam_2_3_content_경변1차, sam_2_average_경변1차, sam_2_stdev_경변1차)
print('유효성분 함량(시료3, 1년차:)', sam_3_1_content_경변1차, sam_3_2_content_경변1차, sam_3_3_content_경변1차, sam_3_average_경변1차, sam_3_stdev_경변1차)
print('1년차 분해율:', 시료1경변분해율1차, 시료2경변분해율1차, 시료3경변분해율1차)

print('유효성분 함량(시료1, 2년차):', sam_1_1_content_경변2차, sam_1_2_content_경변2차, sam_1_3_content_경변2차, sam_1_average_경변2차, sam_1_stdev_경변2차)
print('유효성분 함량(시료2, 2년차):', sam_2_1_content_경변2차, sam_2_2_content_경변2차, sam_2_3_content_경변2차, sam_2_average_경변2차, sam_2_stdev_경변2차)
print('유효성분 함량(시료3, 2년차:)', sam_3_1_content_경변2차, sam_3_2_content_경변2차, sam_3_3_content_경변2차, sam_3_average_경변2차, sam_3_stdev_경변2차)
print('2년차 분해율:', 시료1경변분해율2차, 시료2경변분해율2차, 시료3경변분해율2차)

print('유효성분 함량(시료1, 3년차):', sam_1_1_content_경변3차, sam_1_2_content_경변3차, sam_1_3_content_경변3차, sam_1_average_경변3차, sam_1_stdev_경변3차)
print('유효성분 함량(시료2, 3년차):', sam_2_1_content_경변3차, sam_2_2_content_경변3차, sam_2_3_content_경변3차, sam_2_average_경변3차, sam_2_stdev_경변3차)
print('유효성분 함량(시료3, 3년차:)', sam_3_1_content_경변3차, sam_3_2_content_경변3차, sam_3_3_content_경변3차, sam_3_average_경변3차, sam_3_stdev_경변3차)
print('3년차 분해율:', 시료1경변분해율3차, 시료2경변분해율3차, 시료3경변분해율3차)

print('유효성분 함량(시료1, 4년차):', sam_1_1_content_경변4차, sam_1_2_content_경변4차, sam_1_3_content_경변4차, sam_1_average_경변4차, sam_1_stdev_경변4차)
print('유효성분 함량(시료2, 4년차):', sam_2_1_content_경변4차, sam_2_2_content_경변4차, sam_2_3_content_경변4차, sam_2_average_경변4차, sam_2_stdev_경변4차)
print('유효성분 함량(시료3, 4년차:)', sam_3_1_content_경변4차, sam_3_2_content_경변4차, sam_3_3_content_경변4차, sam_3_average_경변4차, sam_3_stdev_경변4차)
print('4년차 분해율:', 시료1경변분해율4차, 시료2경변분해율4차, 시료3경변분해율4차)

print('저온 안정성 시험 시료1', sam_1_1_content_저온, sam_1_2_content_저온, sam_1_3_content_저온, sam_1_average_저온, sam_1_stdev_저온)
print('저온 안정성 시험 시료2', sam_2_1_content_저온, sam_2_2_content_저온, sam_2_3_content_저온, sam_2_average_저온, sam_2_stdev_저온)
print('저온 안정성 시험 시료3', sam_3_1_content_저온, sam_3_2_content_저온, sam_3_3_content_저온, sam_3_average_저온, sam_3_stdev_저온)
print('저온 분해율:', 시료1경변분해율저온, 시료2경변분해율저온, 시료3경변분해율저온)

print('물리성 ', 검사항목1_경변시작, 검사항목2_경변시작, 검사항목3_경변시작)
print('시험방법 및 조건')
print('약효보증기간 설정')

print('분석기기:', 분석기기1_경변시작, 분석기기2_경변시작, 분석기기3_경변시작)

f = codecs.open("C:\data automation\경변성적서.xml", 'w', 'utf-8')

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
f.write('  <LastAuthor>USER</LastAuthor>\n')
f.write('  <LastPrinted>2016-10-04T04:30:32Z</LastPrinted>\n')
f.write('  <Created>1999-12-11T04:38:33Z</Created>\n')
f.write('  <LastSaved>2016-09-30T02:17:09Z</LastSaved>\n')
f.write('  <Version>14.00</Version>\n')
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
f.write(' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">\n')
f.write('  <AllowPNG/>\n')
f.write(' </OfficeDocumentSettings>\n')
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
f.write('  <Style ss:ID="m179603968">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603988">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604028">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604048">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604068">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604088">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604108">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604128">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604148">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604168">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604188">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604208">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604228">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="11"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179604248">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603708">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603728">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603328">\n')
f.write('   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603348">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603368">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603388">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603408">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603428">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603008">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="9"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603028">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603048">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="9"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603068">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179603088">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602688">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602708">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602828">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602848">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602868">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="9"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602888">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602908">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602928">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="9"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602948">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="9"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602368">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602388">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602408">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602428">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602448">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602468">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602048">\n')
f.write('   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602068">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602088">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602108">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602128">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602148">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602168">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602188">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602208">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601728">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601748">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601768">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601788">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601808">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601828">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <NumberFormat ss:Format="Long Date"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601848">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601868">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601888">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601908">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601928">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601948">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601968">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179601988">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602008">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m179602028">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_ "/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s16">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s17">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s18">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s19">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s20">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s21">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s22">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s23">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s24">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s25">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s26">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s27">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s28">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s29">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s30">\n')
f.write('   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s31">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s32">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s33">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s34">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s35">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s36">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s37">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s38">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s39">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s40">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s41">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s43">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s44">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s45">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s46">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s47">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s48">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s49">\n')
f.write('   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s50">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s51">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="9"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s52">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s53">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s54">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s55">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s56">\n')
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
f.write('  <Style ss:ID="s158">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="8"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s175">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="16"\n')
f.write('    ss:Bold="1"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s176">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="yyyy&quot;년&quot;\ m&quot;월&quot;\ d&quot;일&quot;"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s178">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s188">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('  </Style>\n')
f.write(' </Styles>\n')
f.write(' <Worksheet ss:Name="경변성적서">\n')
f.write('  <Table ss:ExpandedColumnCount="17" ss:ExpandedRowCount="193" x:FullColumns="1"\n')
f.write('   x:FullRows="1" ss:StyleID="s16" ss:DefaultColumnWidth="42"\n')
f.write('   ss:DefaultRowHeight="20.0625">\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="4.5"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="57"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="33"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="21.75"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="33"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="26.25"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="21.75"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="33"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="26.25"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="21.75"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="33"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="26.25"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="21.75"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="33"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="26.25"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="21.75"/>\n')
f.write('   <Column ss:StyleID="s16" ss:AutoFitWidth="0" ss:Width="4.5"/>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="7.5">\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s20"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="14" ss:StyleID="s175"><Data ss:Type="String">경시변화 시험성적서</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="10.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="30.75">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s45"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">시료제조일자&#10;(모집단번호)</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="3" ss:StyleID="s176"><Data ss:Type="String"> ')
f.write(제조년_경변시작 + '년 ' + 제조월_경변시작 + '월 ' + 제조일_경변시작 + '일')
f.write('&#10;(')
f.write(LOTNO_경변시작)
f.write(')</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s29"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="27">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s45"><Data ss:Type="String">분석  책임자&#10;(소속)</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="5" ss:StyleID="s178"><Data ss:Type="String">')
f.write('(주)팜한농 작물보호연구센터')
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s29"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s46"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">성           명</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s27"><Data ss:Type="String">  ')
f.write(책임자)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s29"/>\n')
f.write('    <Cell ss:StyleID="s26"><Data ss:Type="String">인</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('    <Cell ss:StyleID="s29"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s47"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> 시 험 의 뢰 자</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s27"><Data ss:Type="String">')
f.write('(주)팜한농  ')
f.write(의뢰자)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s32"/>\n')
f.write('    <Cell ss:StyleID="s32"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s20"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s39"/>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179603708"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">품    목    명</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="13" ss:MergeDown="1" ss:StyleID="m179603728"><Data\n')
f.write('      ss:Type="String">')
f.write(한글명1_경변시작)

if 한글명2_경변시작 == "":
    pass
else:
    f.write("." + 한글명2_경변시작)

if 한글명3_경변시작 == "":
    pass
else:
    f.write("." + 한글명3_경변시작)

f.write(' ')
f.write(제형분류_경변시작)
f.write('&#10;(')
f.write(시료명1_경변시작)

if 시료명2_경변시작 == "":
    pass
else:
    f.write("." + 시료명2_경변시작)

if 시료명3_경변시작 == "":
    pass
else:
    f.write("." + 시료명3_경변시작)

f.write(')</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="22.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="26.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="s188"><Data ss:Type="String">유효성분의 &#10;명칭 및 함유량</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="13" ss:MergeDown="2" ss:StyleID="m179604028"><ss:Data\n')
f.write('      ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">')
f.write(시료명1IUPAC_경변시작)
f.write('(IUPAC)…………… ')
f.write(함량1_경변시작)

if 시료명2IUPAC_경변시작 == "":
    pass
else:
    f.write('&#10;')

f.write(시료명2IUPAC_경변시작)

if 시료명2IUPAC_경변시작 == "":
    pass
else:
    f.write('(IUPAC)…………… ')
    f.write(함량2_경변시작)
    f.write('&#10;')

if 시료명3IUPAC_경변시작 == "":
    pass
else:
    f.write(시료명3IUPAC_경변시작)
    f.write('&#10;')
    f.write('(IUPAC)…………… ')
    f.write(함량3_경변시작)

f.write('</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="26.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="26.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="28.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s47"><Data ss:Type="String">시험기간</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s36"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> ')
f.write(분석년월일_경변시작 + '~' + list(re.findall(r"(\d+)", 시험기간))[0] + "년" + list(re.findall(r"(\d+)", 시험기간))[1] + "월" + list(re.findall(r"(\d+)", 시험기간))[2] + "일")
f.write('</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179604188"><Data ss:Type="String">포장용기 및 재질</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="5" ss:StyleID="m179604208"><Data ss:Type="String">')
f.write(포장용기_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="16.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="14" ss:StyleID="m179604228"><Data ss:Type="String">시   험   결   과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604248"><Data ss:Type="String">가열 안정성시험</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s40"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179603328"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">분 석 시 기</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="m179603348"><ss:Data\n')
f.write('      ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">1 년 (2주)</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="m179603368"><Data\n')
f.write('      ss:Type="String">2 년 (4주)</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="m179603388"><Data\n')
f.write('      ss:Type="String">3 년 (6주)</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="m179603408"><Data\n')
f.write('      ss:Type="String">4 년 (8주)</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179603428"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">시 작 시</ss:Data></Cell>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s49"><Data ss:Type="String">구분</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179603008"><Data ss:Type="String">유효성분&#10;함량(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179603028"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179603048"><Data ss:Type="String">유효성분&#10;함량(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179603068"><Data ss:Type="String">분해율&#10;(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179603088"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602948"><Data ss:Type="String">유효성분&#10;함량(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602828"><Data ss:Type="String">분해율&#10;(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602848"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602868"><Data ss:Type="String">유효성분&#10;함량(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602888"><Data ss:Type="String">분해율&#10;(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602908"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602928"><Data ss:Type="String">유효성분&#10;함량(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602688"><Data ss:Type="String">분해율&#10;(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="1" ss:StyleID="m179602708"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s50"><Data ss:Type="String">유효성분</Data></Cell>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="33">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s51"><Data ss:Type="String">')
f.write(시료명1_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_1_average_경변시작))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="s158"><Data ss:Type="String">')
f.write(검사항목1_경변시작 + 검사항목2_경변시작 + 검사항목3_경변시작 + 검사항목4_경변시작 + 검사항목5_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_1_average_경변1차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료1경변분해율1차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="s158"><Data ss:Type="String">')
f.write(검사항목1_경변시작 + 검사항목2_경변시작 + 검사항목3_경변시작 + 검사항목4_경변시작 + 검사항목5_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_1_average_경변2차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s55"><Data ss:Type="String">')
f.write(str(시료1경변분해율2차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="s158"><Data ss:Type="String">')
f.write(검사항목1_경변시작 + 검사항목2_경변시작 + 검사항목3_경변시작 + 검사항목4_경변시작 + 검사항목5_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_1_average_경변3차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s55"><Data ss:Type="String">')
f.write(str(시료1경변분해율3차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="s158"><Data ss:Type="String">')
f.write(검사항목1_경변시작 + 검사항목2_경변시작 + 검사항목3_경변시작 + 검사항목4_경변시작 + 검사항목5_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s56"><Data ss:Type="String">')
f.write(str(sam_1_average_경변4차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료1경변분해율4차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="2" ss:StyleID="s158"><Data ss:Type="String">')
f.write(검사항목1_경변시작 + 검사항목2_경변시작 + 검사항목3_경변시작 + 검사항목4_경변시작 + 검사항목5_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="33">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s51"><Data ss:Type="String">')
f.write(시료명2_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_2_average_경변시작))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="5" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_2_average_경변1차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료2경변분해율1차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="8" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_2_average_경변2차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료2경변분해율2차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="11" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_2_average_경변3차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료2경변분해율3차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s56"><Data ss:Type="String">')
f.write(str(sam_2_average_경변4차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료2경변분해율4차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="33">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s51"><Data ss:Type="String">')
f.write(시료명3_경변시작)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_3_average_경변시작))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="5" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_3_average_경변1차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료3경변분해율1차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="8" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_3_average_경변2차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s55"><Data ss:Type="String">')
f.write(str(시료3경변분해율2차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="11" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_3_average_경변3차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s55"><Data ss:Type="String">')
f.write(str(시료3경변분해율3차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="14" ss:StyleID="s54"><Data ss:Type="String">')
f.write(str(sam_3_average_경변4차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"><Data ss:Type="String">')
f.write(str(시료3경변분해율4차))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:Index="17" ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="21.75">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602368"><Data ss:Type="String">시험방법 및 조건</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="6" ss:StyleID="m179602388"><Data ss:Type="String">')

if r3 == None:
    f.write('54℃ 2주 후 분석')
elif r4 == None:
    f.write('54℃ 2주, 4주 후 분석')
elif r5 == None:
    f.write('54℃ 2주, 4주, 6주 후 분석')
else:
    f.write('54℃ 2주, 4주, 6주, 8주 후 분석')

f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179602408"><Data ss:Type="String">약효보증기간 설정</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179602428"><Data ss:Type="String">')

if r3 == None:
    f.write('1년')
elif r4 == None:
    f.write('2년')
elif r5 == None:
    f.write('3년')
else:
    f.write('4년')

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="10.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s39"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s19"/>\n')
f.write('    <Cell ss:StyleID="s20"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="19.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602448"><Data ss:Type="String">저온 안정성시험</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="12" ss:StyleID="m179602468"><Data ss:Type="String">')

if r6 == None:
    f.write('해당사항 없음')
else:
    f.write('0 (±2) ℃ 1주 보관후 분석')

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602048"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">                    분석회수</ss:Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="12" ss:StyleID="m179602068"><Data ss:Type="String">분     석     결     과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602088"><Data ss:Type="String">시작시</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602108"><Data ss:Type="String">1회</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602128"><Data ss:Type="String">2회</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602148"><Data ss:Type="String">3회</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602168"><Data ss:Type="String">평균</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602188"><Data ss:Type="String">분해율(%)</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179602208"><Data ss:Type="String">비  고</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601848"><Data ss:Type="String">')
f.write(시료명1_저온)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601868"><Data ss:Type="String">')
f.write(str(sam_1_1_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601888"><Data ss:Type="String">')
f.write(str(sam_1_2_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601908"><Data ss:Type="String">')
f.write(str(sam_1_3_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601928"><Data ss:Type="String">')
f.write(str(sam_1_average_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601948"><Data ss:Type="String">')
f.write(str(시료1경변분해율저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179601748"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604048"><Data ss:Type="String">')
f.write(시료명2_저온)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604068"><Data ss:Type="String">')
f.write(str(sam_2_1_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604088"><Data ss:Type="String">')
f.write(str(sam_2_2_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604108"><Data ss:Type="String">')
f.write(str(sam_2_3_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604128"><Data ss:Type="String">')
f.write(str(sam_2_average_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179604148"><Data ss:Type="String">')
f.write(str(시료2경변분해율저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179604168"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.25">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601768"><Data ss:Type="String">')
f.write(시료명3_저온)
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601968"><Data ss:Type="String">')
f.write(str(sam_3_1_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179601988"><Data ss:Type="String">')
f.write(str(sam_3_2_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602008"><Data ss:Type="String">')
f.write(str(sam_3_3_content_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179602028"><Data ss:Type="String">')
f.write(str(sam_3_average_저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="m179603968"><Data ss:Type="String">')
f.write(str(시료3경변분해율저온))
f.write('</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m179603988"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s52"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">   첨부 자료</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s35"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="14" ss:StyleID="m179601788"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">        ○ HPLC Chromatograms 첨부</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="14" ss:StyleID="m179601808"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">        ○ 성적계산서</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="6">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s39"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.0625">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="14" ss:StyleID="m179601828"><Data ss:Type="String">')

if 분석년월일_경변4차 == "":
    f.write(분석년월일_경변3차)
elif 분석년월일_경변3차 == "":
    f.write(분석년월일_경변2차)
elif 분석년월일_경변2차 == "":
    f.write(분석년월일_경변1차)
elif 분석년월일_경변1차 == "":
    f.write(분석년월일_경변시작)
else:
    f.write(분석년월일_경변4차)

f.write('</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="10.5">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s39"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s30"/>\n')
f.write('    <Cell ss:StyleID="s30"/>\n')
f.write('    <Cell ss:StyleID="s30"/>\n')
f.write('    <Cell ss:StyleID="s30"/>\n')
f.write('    <Cell ss:StyleID="s30"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="17.0625">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeAcross="14" ss:StyleID="m179601728"><Data ss:Type="String">주식회사 팜한농 작물보호연구센터장   (인)</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="9">\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:StyleID="s44"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:StyleID="s23"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="5.25">\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s53"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('    <Cell ss:StyleID="s22"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0">\n')
f.write('    <Cell ss:Index="2" ss:StyleID="s38"/>\n')
f.write('   </Row>\n')
f.write('  </Table>\n')
f.write('  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n')
f.write('   <PageSetup>\n')
f.write('    <Layout x:CenterHorizontal="1" x:CenterVertical="1"/>\n')
f.write('    <Header x:Margin="0.31496062992125984"/>\n')
f.write('    <Footer x:Margin="0.31496062992125984"/>\n')
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
f.write('     <ActiveRow>12</ActiveRow>\n')
f.write('     <ActiveCol>23</ActiveCol>\n')
f.write('    </Pane>\n')
f.write('   </Panes>\n')
f.write('   <ProtectObjects>False</ProtectObjects>\n')
f.write('   <ProtectScenarios>False</ProtectScenarios>\n')
f.write('  </WorksheetOptions>\n')
f.write(' </Worksheet>\n')
f.write('</Workbook>\n')
f.close()

f = codecs.open("C:\data automation\경시변화시험내용.xml", 'w', 'utf-8')

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
f.write('  <LastAuthor>USER</LastAuthor>\n')
f.write('  <LastPrinted>2016-09-30T02:25:22Z</LastPrinted>\n')
f.write('  <Created>1999-12-11T04:38:33Z</Created>\n')
f.write('  <LastSaved>2016-11-29T08:21:30Z</LastSaved>\n')
f.write('  <Version>14.00</Version>\n')
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
f.write(' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">\n')
f.write('  <AllowPNG/>\n')
f.write(' </OfficeDocumentSettings>\n')
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
f.write('  <Style ss:ID="m190154944">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154964">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190155044">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154644">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154684">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154704">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154764">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154784">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154804">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154364">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154004">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154044">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154084">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154124">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154164">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190154204">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153684">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153724">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153764">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153804">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153844">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153884">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153924">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153364">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153404">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153444">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153464">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153484">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153504">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153524">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153024">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153044">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153104">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153144">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153184">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153284">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153304">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="m190153324">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s16">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s17">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s18">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="Long Date"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s19">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s20">\n')
f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s21">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s22">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s23">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s24">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s25">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s26">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s27">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s28">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0_);[Red]\(0\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s29">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s30">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s31">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s32">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s33">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s34">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s35">\n')
f.write('   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s36">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s37">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s38">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s39">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s41">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s42">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00000_);[Red]\(0.00000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s43">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.0000_);[Red]\(0.0000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s44">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.0000_);[Red]\(0.0000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s45">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s46">\n')
f.write('   <Alignment ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s47">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s48">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00000_);[Red]\(0.00000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s77">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.00_);[Red]\(0.00\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s80">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s88">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders>\n')
f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
f.write('   </Borders>\n')
f.write('   <Interior/>\n')
f.write('   <NumberFormat ss:Format="0.000_);[Red]\(0.000\)"/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s92">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write('  <Style ss:ID="s93">\n')
f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
f.write('   <Borders/>\n')
f.write('   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="14"\n')
f.write('    ss:Bold="1" ss:Underline="Single"/>\n')
f.write('   <Interior/>\n')
f.write('  </Style>\n')
f.write(' </Styles>\n')
f.write(' <Worksheet ss:Name="경시변화시험내용">\n')
f.write('  <Table ss:ExpandedColumnCount="15" ss:ExpandedRowCount="93" x:FullColumns="1"\n')
f.write('   x:FullRows="1" ss:StyleID="s19" ss:DefaultColumnWidth="42"\n')
f.write('   ss:DefaultRowHeight="12">\n')
f.write('   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="4.5"/>\n')
f.write('   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="50.25"/>\n')
f.write('   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="61.5"/>\n')
f.write('   <Column ss:StyleID="s19" ss:Width="42.75"/>\n')
f.write('   <Column ss:Index="9" ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="55.5"/>\n')
f.write('   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="48.75"/>\n')
f.write('   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="4.5"/>\n')
f.write('   <Column ss:Index="13" ss:StyleID="s19" ss:Width="45"/>\n')
f.write('   <Column ss:StyleID="s19" ss:Width="45.75" ss:Span="1"/>\n')
f.write('   <Row>\n')
f.write('    <Cell ss:StyleID="s30"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s31"/>\n')
f.write('    <Cell ss:StyleID="s32"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="9.75">\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:MergeAcross="8" ss:StyleID="s92"/>\n')
f.write('    <Cell ss:StyleID="s16"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:Height="18.75">\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:MergeAcross="8" ss:StyleID="s93"><Data ss:Type="String">아짐설퓨론.벤조비사이클론.옥사지클로메폰 수면부상성입제의 경시변화시험내용</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s16"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="13.5">\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:MergeAcross="8" ss:StyleID="s92"/>\n')
f.write('    <Cell ss:StyleID="s16"/>\n')
f.write('   </Row>\n')
f.write('   <Row>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s20"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">○ 시험기간 :<B>  </B><Font>20</Font>14년 2월 12일 ~ 4월 9일</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s35"><Data ss:Type="String">○ 공시품수량</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s17"><Data ss:Type="String">: 1회 분석 : 150g× 3봉</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s16"/>\n')
f.write('   </Row>\n')
f.write('   <Row>\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s36"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">○ 시험내용 : 54℃ 2주, 4주, 6주, 8주 후 분석</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"><Data ss:Type="String">: 총시료량 : 150g×15봉</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s16"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="13.5">\n')
f.write('    <Cell ss:StyleID="s33"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s17"/>\n')
f.write('    <Cell ss:StyleID="s16"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s38"><Data ss:Type="String">분석일자</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="7" ss:StyleID="m190154364"><Data ss:Type="String">검사결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s21"/>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="s77"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">유효성분(%)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">1번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">2번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">3번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"><Data ss:Type="String">주성분평균</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="String">분해율</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">0.51100000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.51</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.51200000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.51100000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">7.0090000000000003</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">7.0140000000000002</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">7.0170000000000003</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">7.0129999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">2.7869999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.778</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7879999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7839999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"><Data ss:Type="String">2014. 2. 12</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="5" ss:StyleID="m190154704"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="s34"><Data ss:Type="String">항 목</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154644"><Data ss:Type="String">검사결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"><Data ss:Type="String">(시작)</Data></Cell>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목1</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154684"><Data ss:Type="String">검사1결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목2</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154764"><Data ss:Type="String">검사2결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목3</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154784"><Data ss:Type="String">검사3결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목4</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154804"><Data ss:Type="String">검사4결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목5</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154964"><Data ss:Type="String">검사5결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s27"/>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="s77"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">유효성분(%)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">1번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">2번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">3번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"><Data ss:Type="String">주성분평균</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="String">분해율</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">0.502</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.503</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.505</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.503</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">1.57</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">6.9909999999999997</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9779999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9859999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9850000000000003</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">0.4</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">2.7679999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7639999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7709999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7679999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">0.56999999999999995</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"><Data ss:Type="String">2014. 2. 26</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="5" ss:StyleID="m190154004"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="s34"><Data ss:Type="String">항 목</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154044"><Data ss:Type="String">검사결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"><Data ss:Type="String">(1 년차)</Data></Cell>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목1</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190155044"><Data ss:Type="String">검사1결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목2</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154124"><Data ss:Type="String">검사2결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목3</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154164"><Data ss:Type="String">검사3결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목4</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154204"><Data ss:Type="String">검사4결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s26"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목5</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154084"><Data ss:Type="String">검사5결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s27"/>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="s77"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">유효성분(%)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">1번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">2번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">3번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"><Data ss:Type="String">주성분평균</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="String">분해율</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">0.496</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.495</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.49399999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.495</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">3.13</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">6.9669999999999996</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.952</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.96</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.96</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">0.76</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">2.7549999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7509999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7639999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7570000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">0.97</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"><Data ss:Type="String">2014. 3. 12</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="5" ss:StyleID="m190153684"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="s34"><Data ss:Type="String">항 목</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153724"><Data ss:Type="String">검사결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"><Data ss:Type="String">(2 년차)</Data></Cell>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목1</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153804"><Data ss:Type="String">검사1결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목2</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153844"><Data ss:Type="String">검사2결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목3</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153884"><Data ss:Type="String">검사3결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목4</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153924"><Data ss:Type="String">검사4결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('    <Cell ss:StyleID="s28"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s26"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목5</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153764"><Data ss:Type="String">검사5결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('    <Cell ss:StyleID="s29"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s27"/>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="s77"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">유효성분(%)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">1번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">2번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">3번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"><Data ss:Type="String">주성분평균</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="String">분해율</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">0.48499999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.48699999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.48599999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.48599999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">4.8899999999999997</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">6.94</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9240000000000004</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9409999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9349999999999996</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">1.1100000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">2.7490000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7440000000000002</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7330000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.742</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">1.51</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"><Data ss:Type="String">2014. 3. 26</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="5" ss:StyleID="m190153364"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="s34"><Data ss:Type="String">항 목</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153404"><Data ss:Type="String">검사결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"><Data ss:Type="String">(3 년차)</Data></Cell>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목1</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153464"><Data ss:Type="String">검사1결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목2</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153484"><Data ss:Type="String">검사2결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목3</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153504"><Data ss:Type="String">검사3결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목4</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153524"><Data ss:Type="String">검사4결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s26"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목5</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153444"><Data ss:Type="String">검사5결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="s77"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">유효성분(%)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">1번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">2번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">3번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"><Data ss:Type="String">주성분평균</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="String">분해율</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">0.47799999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.47699999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.48</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.47799999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">6.46</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">6.9119999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9180000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9080000000000004</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9130000000000003</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">1.43</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">2.73</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.722</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7330000000000001</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7280000000000002</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="Number">2.0099999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s44"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s18"><Data ss:Type="String">2014. 4. 9</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="5" ss:StyleID="m190153104"><Data ss:Type="String">물리성</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="1" ss:StyleID="s34"><Data ss:Type="String">항 목</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153144"><Data ss:Type="String">검사결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">(4 년차)</ss:Data></Cell>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목1</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153284"><Data ss:Type="String">검사1결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목2</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153304"><Data ss:Type="String">검사2결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목3</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153324"><Data ss:Type="String">검사3결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목4</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190154944"><Data ss:Type="String">검사4결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:StyleID="s25"/>\n')
f.write('    <Cell ss:Index="4" ss:MergeAcross="1" ss:StyleID="s88"><Data ss:Type="String">검사항목5</Data></Cell>\n')
f.write('    <Cell ss:MergeAcross="4" ss:StyleID="m190153184"><Data ss:Type="String">검사5결과</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="m190153024"><Data ss:Type="String">평 균</Data></Cell>\n')
f.write('    <Cell ss:MergeDown="3" ss:StyleID="m190153044"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">유효성분(%)</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">1번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><Data ss:Type="String">2번</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40">3번 </ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s22"><ss:Data ss:Type="String"\n')
f.write('      xmlns="http://www.w3.org/TR/REC-html40"> -</ss:Data></Cell>\n')
f.write('    <Cell ss:StyleID="s23"><Data ss:Type="String">주성분평균</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"><Data ss:Type="String">분해율</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">0.49399999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.49399999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.495</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">0.49399999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">6.9640000000000004</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9569999999999999</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9619999999999997</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">6.9610000000000003</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s37"/>\n')
f.write('    <Cell ss:Index="4" ss:StyleID="s34"><Data ss:Type="Number">2.758</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7519999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.758</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s24"/>\n')
f.write('    <Cell ss:StyleID="s34"><Data ss:Type="Number">2.7559999999999998</Data></Cell>\n')
f.write('    <Cell ss:StyleID="s43"/>\n')
f.write('    <Cell ss:StyleID="s41"/>\n')
f.write('    <Cell ss:Index="13" ss:StyleID="s42"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:AutoFitHeight="0" ss:Height="12.75" ss:StyleID="s39">\n')
f.write('    <Cell ss:StyleID="s45"/>\n')
f.write('    <Cell ss:MergeAcross="8" ss:StyleID="s80"/>\n')
f.write('    <Cell ss:StyleID="s46"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s48"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39">\n')
f.write('    <Cell ss:Index="3" ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('    <Cell ss:StyleID="s47"/>\n')
f.write('   </Row>\n')
f.write('   <Row ss:StyleID="s39" ss:Span="20"/>\n')
f.write('  </Table>\n')
f.write('  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n')
f.write('   <PageSetup>\n')
f.write('    <Layout x:CenterHorizontal="1" x:CenterVertical="1"/>\n')
f.write('    <Header x:Margin="0.31496062992125984"/>\n')
f.write('    <Footer x:Margin="0.31496062992125984"/>\n')
f.write('    <PageMargins x:Bottom="0" x:Left="0" x:Right="0" x:Top="0"/>\n')
f.write('   </PageSetup>\n')
f.write('   <Print>\n')
f.write('    <ValidPrinterInfo/>\n')
f.write('    <PaperSizeIndex>9</PaperSizeIndex>\n')
f.write('    <HorizontalResolution>600</HorizontalResolution>\n')
f.write('    <VerticalResolution>600</VerticalResolution>\n')
f.write('   </Print>\n')
f.write('   <Selected/>\n')
f.write('   <TopRowVisible>15</TopRowVisible>\n')
f.write('   <Panes>\n')
f.write('    <Pane>\n')
f.write('     <Number>3</Number>\n')
f.write('     <ActiveRow>51</ActiveRow>\n')
f.write('     <ActiveCol>13</ActiveCol>\n')
f.write('    </Pane>\n')
f.write('   </Panes>\n')
f.write('   <ProtectObjects>False</ProtectObjects>\n')
f.write('   <ProtectScenarios>False</ProtectScenarios>\n')
f.write('  </WorksheetOptions>\n')
f.write(' </Worksheet>\n')
f.write('</Workbook>\n')
f.close()













