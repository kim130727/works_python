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
시험번호 = r1[1]
시험책임자 = r1[7]
시험담당자 = r1[8]
시험의뢰자 = r1[6]

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
    검사항목1결과 = r1[77]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
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
    포장용기_경변시작 = "합성수지병, HDPE재질"
    검사항목1_경변시작 = "수화성: " + r1[78]
    검사항목2_경변시작 = "분말도: " + r1[81]
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    포장용기단위_경변시작 = "병"
    수량 = "500ml"
    검사항목1결과 = r1[78]
    검사항목2결과 = r1[81]
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""

elif 제형코드_경변시작 == "WP":
    제형분류_경변시작 = "수화제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
    검사항목1_경변시작 = "수화성"
    검사항목2_경변시작 = "분말도"
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    검사항목1결과 = r1[78]
    검사항목2결과 = r1[81]
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "은박코팅봉투, 알루미늄재질"

elif 제형코드_경변시작 == "SP":
    제형분류_경변시작 = "수용제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
elif 제형코드_경변시작 == "OD":
    제형분류_경변시작 = "유상수화제"
    수량_경변시작 = "500ml"
    검사항목1_경변시작 = "유화성"
    검사항목2_경변시작 = "수화성"
    검사항목3_경변시작 = "분말도"
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    검사항목1결과_경변시작 = r1[77]
    검사항목2결과_경변시작 = r1[78]
    검사항목3결과_경변시작 = r1[81]
    검사항목4결과_경변시작 = ""
    검사항목5결과_경변시작 = ""
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
    검사항목1결과 = r1[81]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""

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
    검사항목1_경변시작 = ""
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    검사항목1결과 = ""
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기 = "은박코팅봉투, 알루미늄재질"

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
    검사항목1결과 = r1[94]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""

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
    검사항목1결과 = r1[78]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
    포장용기_경변시작 = "합성수지병, PE재질"

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
    검사항목1결과 = r1[77]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""

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
    검사항목1결과 = r1[77]
    검사항목2결과 = r1[78]
    검사항목3결과 = r1[81]
    검사항목4결과 = ""
    검사항목5결과 = ""

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
    검사항목1결과 = r1[79]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""

elif 제형코드_경변시작 == "SO":
    제형분류_경변시작 = "수면전개제"
    수량_경변시작 = "500ml"
    포장용기_경변시작 = "합성수지병, PE재질"
    포장용기단위_경변시작 = "병"
    검사항목1_경변시작 = ""
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    검사항목1결과 = ""
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
elif 제형코드_경변시작 == "WS":
    제형분류_경변시작 = "종자처리수화제"
    수량_경변시작 = "500g"
    포장용기_경변시작 = "은박코팅봉투, 알루미늄재질"
    검사항목1_경변시작 = "수화성"
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    검사항목1결과 = r1[78]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
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
    검사항목1_경변시작 = "수화성"
    검사항목2_경변시작 = ""
    검사항목3_경변시작 = ""
    검사항목4_경변시작 = ""
    검사항목5_경변시작 = ""
    검사항목1결과 = r1[78]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
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
    검사항목1결과 = r1[93]
    검사항목2결과 = ""
    검사항목3결과 = ""
    검사항목4결과 = ""
    검사항목5결과 = ""
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

factor_std1_경변시작 = float(std_g1_경변시작) * float(std_content1_경변시작) * float(std_IS_area1_경변시작) / float(std_AI_area1_경변시작)
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
    std_content2_경변시작 = ""
    std_g2_경변시작 = ""
    sam_2_1_g_경변시작 =""
    sam_2_2_g_경변시작 =""
    sam_2_3_g_경변시작 =""

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

    factor_std2_경변시작 = float(std_g2_경변시작) * float(std_content2_경변시작) * float(std_IS_area2_경변시작) / float(std_AI_area2_경변시작)
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
    std_content3_경변시작 = ""
    std_g3_경변시작 = ""
    sam_3_1_g_경변시작 =""
    sam_3_2_g_경변시작 =""
    sam_3_3_g_경변시작 =""

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

    factor_std3_경변시작 = float(std_g3_경변시작) * float(std_content3_경변시작) * float(std_IS_area3_경변시작) / float(std_AI_area3_경변시작)
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

    factor_std1_경변1차 = float(std_g1_경변1차) * float(std_content1_경변1차) * float(std_IS_area1_경변1차) / float(std_AI_area1_경변1차)
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
        std_content2_경변1차 =""
        std_g2_경변1차 =""
        sam_2_1_g_경변1차 =""
        sam_2_2_g_경변1차 =""
        sam_2_3_g_경변1차 =""


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

        factor_std2_경변1차 = float(std_g2_경변1차) * float(std_content2_경변1차) * float(std_IS_area2_경변1차) / float(std_AI_area2_경변1차)
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
        std_content3_경변1차 =""
        std_g3_경변1차 =""
        sam_3_1_g_경변1차 =""
        sam_3_2_g_경변1차 =""
        sam_3_3_g_경변1차 =""

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

        factor_std3_경변1차 = float(std_g3_경변1차) * float(std_content3_경변1차) * float(std_IS_area3_경변1차) / float(std_AI_area3_경변1차)
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

    factor_std1_경변2차 = float(std_g1_경변2차) * float(std_content1_경변2차) * float(std_IS_area1_경변2차) / float(std_AI_area1_경변2차)
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
        std_content2_경변2차 =""
        std_g2_경변2차 =""
        sam_2_1_g_경변2차 =""
        sam_2_2_g_경변2차 =""
        sam_2_3_g_경변2차 =""

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

        factor_std2_경변2차 = float(std_g2_경변2차) * float(std_content2_경변2차) * float(std_IS_area2_경변2차) / float(std_AI_area2_경변2차)
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
        std_content3_경변2차 =""
        std_g3_경변2차 =""
        sam_3_1_g_경변2차 =""
        sam_3_2_g_경변2차 =""
        sam_3_3_g_경변2차 =""

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

        factor_std3_경변2차 = float(std_g3_경변2차) * float(std_content3_경변2차) * float(std_IS_area3_경변2차) / float(std_AI_area3_경변2차)
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

    factor_std1_경변3차 = float(std_g1_경변3차) * float(std_content1_경변3차) * float(std_IS_area1_경변3차) / float(std_AI_area1_경변3차)
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
        std_content2_경변3차 =""
        std_g2_경변3차 =""
        sam_2_1_g_경변3차 =""
        sam_2_2_g_경변3차 =""
        sam_2_3_g_경변3차 =""

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

        factor_std2_경변3차 = float(std_g2_경변3차) * float(std_content2_경변3차) * float(std_IS_area2_경변3차) / float(std_AI_area2_경변3차)
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
        std_content3_경변3차 =""
        std_g3_경변3차 =""
        sam_3_1_g_경변3차 =""
        sam_3_2_g_경변3차 =""
        sam_3_3_g_경변3차 =""

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

        factor_std3_경변3차 = float(std_g3_경변3차) * float(std_content3_경변3차) * float(std_IS_area3_경변3차) / float(std_AI_area3_경변3차)
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
    std_content1_경변4차 = ""
    std_g1_경변4차 = ""
    sam_1_1_g_경변4차 = ""
    sam_1_2_g_경변4차 = ""
    sam_1_3_g_경변4차 = ""
    std_content2_경변4차 = ""
    std_g2_경변4차 = ""
    sam_2_1_g_경변4차 = ""
    sam_2_2_g_경변4차 = ""
    sam_2_3_g_경변4차 = ""
    std_content3_경변4차 = ""
    std_g3_경변4차 = ""
    sam_3_1_g_경변4차 = ""
    sam_3_2_g_경변4차 = ""
    sam_3_3_g_경변4차 = ""

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

    factor_std1_경변4차 = float(std_g1_경변4차) * float(std_content1_경변4차) * float(std_IS_area1_경변4차) / float(std_AI_area1_경변4차)
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
        std_content2_경변4차 =""
        std_g2_경변4차 =""
        sam_2_1_g_경변4차 =""
        sam_2_2_g_경변4차 =""
        sam_2_3_g_경변4차 =""

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

        factor_std2_경변4차 = float(std_g2_경변4차) * float(std_content2_경변4차) * float(std_IS_area2_경변4차) / float(std_AI_area2_경변4차)
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
        std_content3_경변4차 =""
        std_g3_경변4차 =""
        sam_3_1_g_경변4차 =""
        sam_3_2_g_경변4차 =""
        sam_3_3_g_경변4차 =""

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

        factor_std3_경변4차 = float(std_g3_경변4차) * float(std_content3_경변4차) * float(std_IS_area3_경변4차) / float(std_AI_area3_경변4차)
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
    std_content1_저온 = ""
    std_content2_저온 = ""
    std_content3_저온 = ""
    std_g1_저온 = ""
    sam_1_1_g_저온= ""
    sam_1_2_g_저온 = ""
    sam_1_3_g_저온 = ""
    std_g2_저온 = ""
    sam_2_1_g_저온= ""
    sam_2_2_g_저온 = ""
    sam_2_3_g_저온 = ""
    std_g3_저온 = ""
    sam_3_1_g_저온= ""
    sam_3_2_g_저온 = ""
    sam_3_3_g_저온 = ""

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

    factor_std1_저온 = float(std_g1_저온) * float(std_content1_저온) * float(std_IS_area1_저온) / float(std_AI_area1_저온)
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
        std_content2_저온 =""
        std_g2_저온 =""
        sam_2_1_g_저온 =""
        sam_2_2_g_저온 =""
        sam_2_3_g_저온 =""

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

        factor_std2_저온 = float(std_g2_저온) * float(std_content2_저온) * float(std_IS_area2_저온) / float(std_AI_area2_저온)
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
        std_content3_저온 =""
        std_g3_저온 =""
        sam_3_1_g_저온 =""
        sam_3_2_g_저온 =""
        sam_3_3_g_저온 =""

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

        factor_std3_저온 = float(std_g3_저온) * float(std_content3_저온) * float(std_IS_area3_저온) / float(std_AI_area3_저온)
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

print ("********************************************************************************************")

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

print("********************************************************************************************")
print("농약품목의 경시변화시험 보고서")
print(한글명1_경변시작, 한글명2_경변시작, 한글명3_경변시작, 제형분류_경변시작, " 경시변화시험")
print("시험번호 : ", 시험번호)
print("(주)팜한농 작물보호연구센터")
print("********************************************************************************************")
print("제출문")
print("제목 :",시험물질_경변시작,"의 경시변화시험")
print("시험번호 : ", 시험번호)
print("본 시험에 사용된 기준은 다음과 같습니다.")
print("1. 농촌진흥청 고시 농약 및 원제의 등록기준 및 농약의 검사방법 및 부정불량 농약 처리요령")
print("2. 농촌진흥청 고시 농약 등의 시험연구기고나 지정 및 관리기준")
print("본 보고서에 기술된 시험과정은 시험책임자의 책임 하에 수행되었으며,")
print("위의 기준을 준수하여 실시하였으며, 시험결과는 생성된 모든 시험기초자료를 토대로 작성되었습니다.")
print("시험 책임자 : ",시험책임자)
print("********************************************************************************************")
print("농약품목의 경시변화시험 성적서")
print("시험번호 ", 시험번호, ", 시험분야 : 이화학, 시험년도 :",list(re.findall(r"(\d+)", 시험기간))[0])
print("시험항목 ",시험물질_경변시작,"의 경시변화시험")
print('시험기간: ', 분석년월일_경변시작, '~', list(re.findall(r"(\d+)", 시험기간))[0] + "년 ", list(re.findall(r"(\d+)", 시험기간))[1] + "월 ",
      list(re.findall(r"(\d+)", 시험기간))[2] + "일")
print("시험기관: (주)팜한농 작물보호연구센터")
print("시험담당자: ", 시험담당자)
print("시험의뢰자: ", 시험의뢰자)
print("")
print("1. 목적")
print (시험물질_경변시작,"의 경과시간별 유효성분 등의 경시안정성을 구명하여 농약품목등록의")
print ("이화학성 평가 및 품질관리 기준 설정 등을 위한 기초자료로 활용코자 함")
print ("2. 시험방법")
print ("가. 시험약제 :", 시험물질_경변시작)
print ("시험물질정보 :", LOTNO_경변시작, ' 제조, 제조일자:', 제조년_경변시작 + '년', 제조월_경변시작 + '월', 제조일_경변시작 + '일')
print ("나. 시험세부항목")
print ("유효성분의 경시 안정성")
print ("가열안정성")
print ("시료의 경시적 제제안정성")
print ("물리성 :", 검사항목1_경변시작, 검사항목2_경변시작, 검사항목3_경변시작, 검사항목4_경변시작, 검사항목5_경변시작)
print ("다. 가열안정성 시험방법")
print ("시료를 시험0일차 초기 분석하고, 실온에서 보관중인 동일 모집단 공시품을 2주 경과시마다")
print ("항온기에 투입한 후 경시변화 시험 종료일에 각 주차별(2,4,6,8주차) 시료를")
print ("함께 꺼내어 일괄 분석한다.")
print ("")
print ("3. 가열안정성 시험성적")
print ("시작시(0일차) 시험결과 (분석일 : ",분석일_경변시작," )")
print ("유효성분 :", 시료명1_경변시작)
print ("표준품 순도(%) ", std_content1_경변시작, "표준품 무게 ",std_g1_경변시작,"시료 무게 ",sam_1_1_g_경변시작," " ,sam_1_2_g_경변시작," ",sam_1_3_g_경변시작)
print('유효성분 함량(시료1, 시작):', sam_1_1_content_경변시작, sam_1_2_content_경변시작, sam_1_3_content_경변시작, sam_1_average_경변시작)
print ("유효성분 :", 시료명2_경변시작)
print ("표준품 순도(%) ", std_content2_경변시작, "표준품 무게 ",std_g2_경변시작,"시료 무게 ",sam_2_1_g_경변시작," " ,sam_2_2_g_경변시작," ",sam_2_3_g_경변시작)
print('유효성분 함량(시료2, 시작):', sam_2_1_content_경변시작, sam_2_2_content_경변시작, sam_2_3_content_경변시작, sam_2_average_경변시작)
print ("유효성분 :", 시료명3_경변시작)
print ("표준품 순도(%) ", std_content3_경변시작, "표준품 무게 ",std_g3_경변시작,"시료 무게 ",sam_3_1_g_경변시작," " ,sam_3_2_g_경변시작," ",sam_3_3_g_경변시작)
print('유효성분 함량(시료3, 시작):', sam_3_1_content_경변시작, sam_3_2_content_경변시작, sam_3_3_content_경변시작, sam_3_average_경변시작)

print ("2주차 시험결과 (분석일 : ",분석일_경변1차," )")
print ("유효성분 :", 시료명1_경변1차)
print ("표준품 순도(%) ", std_content1_경변1차, "표준품 무게 ",std_g1_경변1차,"시료 무게 ",sam_1_1_g_경변1차," " ,sam_1_2_g_경변1차," ",sam_1_3_g_경변1차)
print('유효성분 함량(시료1, 1년차):', sam_1_1_content_경변1차, sam_1_2_content_경변1차, sam_1_3_content_경변1차, sam_1_average_경변1차)
print ("유효성분 :", 시료명2_경변1차)
print ("표준품 순도(%) ", std_content2_경변1차, "표준품 무게 ",std_g2_경변1차,"시료 무게 ",sam_2_1_g_경변1차," " ,sam_2_2_g_경변1차," ",sam_2_3_g_경변1차)
print('유효성분 함량(시료2, 1년차):', sam_2_1_content_경변1차, sam_2_2_content_경변1차, sam_2_3_content_경변1차, sam_2_average_경변1차)
print ("유효성분 :", 시료명3_경변1차)
print ("표준품 순도(%) ", std_content3_경변1차, "표준품 무게 ",std_g3_경변1차,"시료 무게 ",sam_3_1_g_경변1차," " ,sam_3_2_g_경변1차," ",sam_3_3_g_경변1차)
print('유효성분 함량(시료3, 1년차:)', sam_3_1_content_경변1차, sam_3_2_content_경변1차, sam_3_3_content_경변1차, sam_3_average_경변1차)
print('1년차 분해율:', 시료1경변분해율1차, 시료2경변분해율1차, 시료3경변분해율1차)

print ("4주차 시험결과 (분석일 : ",분석일_경변2차," )")
print ("유효성분 :", 시료명1_경변2차)
print ("표준품 순도(%) ", std_content1_경변2차, "표준품 무게 ",std_g1_경변2차,"시료 무게 ",sam_1_1_g_경변2차," " ,sam_1_2_g_경변2차," ",sam_1_3_g_경변2차)
print('유효성분 함량(시료1, 2년차):', sam_1_1_content_경변2차, sam_1_2_content_경변2차, sam_1_3_content_경변2차, sam_1_average_경변2차)
print ("유효성분 :", 시료명2_경변2차)
print ("표준품 순도(%) ", std_content2_경변2차, "표준품 무게 ",std_g2_경변2차,"시료 무게 ",sam_2_1_g_경변2차," " ,sam_2_2_g_경변2차," ",sam_2_3_g_경변2차)
print('유효성분 함량(시료2, 2년차):', sam_2_1_content_경변2차, sam_2_2_content_경변2차, sam_2_3_content_경변2차, sam_2_average_경변2차)
print ("유효성분 :", 시료명3_경변2차)
print ("표준품 순도(%) ", std_content3_경변2차, "표준품 무게 ",std_g3_경변2차,"시료 무게 ",sam_3_1_g_경변2차," " ,sam_3_2_g_경변2차," ",sam_3_3_g_경변2차)
print('유효성분 함량(시료3, 2년차:)', sam_3_1_content_경변2차, sam_3_2_content_경변2차, sam_3_3_content_경변2차, sam_3_average_경변2차)
print('2년차 분해율:', 시료1경변분해율2차, 시료2경변분해율2차, 시료3경변분해율2차)

print ("6주차 시험결과 (분석일 : ",분석일_경변3차," )")
print ("유효성분 :", 시료명1_경변3차)
print ("표준품 순도(%) ", std_content1_경변3차, "표준품 무게 ",std_g1_경변3차,"시료 무게 ",sam_1_1_g_경변3차," " ,sam_1_2_g_경변3차," ",sam_1_3_g_경변3차)
print('유효성분 함량(시료1, 3년차):', sam_1_1_content_경변3차, sam_1_2_content_경변3차, sam_1_3_content_경변3차, sam_1_average_경변3차)
print ("유효성분 :", 시료명2_경변3차)
print ("표준품 순도(%) ", std_content2_경변3차, "표준품 무게 ",std_g2_경변3차,"시료 무게 ",sam_2_1_g_경변3차," " ,sam_2_2_g_경변3차," ",sam_2_3_g_경변3차)
print('유효성분 함량(시료2, 3년차):', sam_2_1_content_경변3차, sam_2_2_content_경변3차, sam_2_3_content_경변3차, sam_2_average_경변3차)
print ("유효성분 :", 시료명3_경변3차)
print ("표준품 순도(%) ", std_content3_경변3차, "표준품 무게 ",std_g3_경변3차,"시료 무게 ",sam_3_1_g_경변3차," " ,sam_3_2_g_경변3차," ",sam_3_3_g_경변3차)
print('유효성분 함량(시료3, 3년차:)', sam_3_1_content_경변3차, sam_3_2_content_경변3차, sam_3_3_content_경변3차, sam_3_average_경변3차)
print('3년차 분해율:', 시료1경변분해율3차, 시료2경변분해율3차, 시료3경변분해율3차)

print ("8주차 시험결과 (분석일 : ",분석일_경변4차," )")
print ("유효성분 :", 시료명1_경변4차)
print ("표준품 순도(%) ", std_content1_경변4차, "표준품 무게 ",std_g1_경변4차,"시료 무게 ",sam_1_1_g_경변4차," " ,sam_1_2_g_경변4차," ",sam_1_3_g_경변4차)
print('유효성분 함량(시료1, 4년차):', sam_1_1_content_경변4차, sam_1_2_content_경변4차, sam_1_3_content_경변4차, sam_1_average_경변4차)
print ("유효성분 :", 시료명2_경변4차)
print ("표준품 순도(%) ", std_content2_경변4차, "표준품 무게 ",std_g2_경변4차,"시료 무게 ",sam_2_1_g_경변4차," " ,sam_2_2_g_경변4차," ",sam_2_3_g_경변4차)
print('유효성분 함량(시료2, 4년차):', sam_2_1_content_경변4차, sam_2_2_content_경변4차, sam_2_3_content_경변4차, sam_2_average_경변4차)
print ("유효성분 :", 시료명3_경변4차)
print ("표준품 순도(%) ", std_content3_경변4차, "표준품 무게 ",std_g3_경변4차,"시료 무게 ",sam_3_1_g_경변4차," " ,sam_3_2_g_경변4차," ",sam_3_3_g_경변4차)
print('유효성분 함량(시료3, 4년차:)', sam_3_1_content_경변4차, sam_3_2_content_경변4차, sam_3_3_content_경변4차, sam_3_average_경변4차)
print('4년차 분해율:', 시료1경변분해율4차, 시료2경변분해율4차, 시료3경변분해율4차)

print ("4. 저온안정성 시험성적")

print ("시작시(0일차) 시험결과 (분석일 : ",분석일_경변시작," )")
print ("유효성분 :", 시료명1_경변시작)
print ("표준품 순도(%) ", std_content1_경변시작, "표준품 무게 ",std_g1_경변시작,"시료 무게 ",sam_1_1_g_경변시작," " ,sam_1_2_g_경변시작," ",sam_1_3_g_경변시작)
print('유효성분 함량(시료1, 시작):', sam_1_1_content_경변시작, sam_1_2_content_경변시작, sam_1_3_content_경변시작, sam_1_average_경변시작)
print ("유효성분 :", 시료명2_경변시작)
print ("표준품 순도(%) ", std_content2_경변시작, "표준품 무게 ",std_g2_경변시작,"시료 무게 ",sam_2_1_g_경변시작," " ,sam_2_2_g_경변시작," ",sam_2_3_g_경변시작)
print('유효성분 함량(시료2, 시작):', sam_2_1_content_경변시작, sam_2_2_content_경변시작, sam_2_3_content_경변시작, sam_2_average_경변시작)
print ("유효성분 :", 시료명3_경변시작)
print ("표준품 순도(%) ", std_content3_경변시작, "표준품 무게 ",std_g3_경변시작,"시료 무게 ",sam_3_1_g_경변시작," " ,sam_3_2_g_경변시작," ",sam_3_3_g_경변시작)
print('유효성분 함량(시료3, 시작):', sam_3_1_content_경변시작, sam_3_2_content_경변시작, sam_3_3_content_경변시작, sam_3_average_경변시작)

print ("7일차 시험결과 (분석일 : ",분석일_저온," )")
print ("유효성분 :", 시료명1_저온)
print ("표준품 순도(%) ", std_content1_저온, "표준품 무게 ",std_g1_저온,"시료 무게 ",sam_1_1_g_저온," " ,sam_1_2_g_저온," ",sam_1_3_g_저온)
print('저온 안정성 시험 시료1', sam_1_1_content_저온, sam_1_2_content_저온, sam_1_3_content_저온, sam_1_average_저온)
print ("유효성분 :", 시료명2_저온)
print ("표준품 순도(%) ", std_content2_저온, "표준품 무게 ",std_g2_저온,"시료 무게 ",sam_2_1_g_저온," " ,sam_2_2_g_저온," ",sam_2_3_g_저온)
print('저온 안정성 시험 시료2', sam_2_1_content_저온, sam_2_2_content_저온, sam_2_3_content_저온, sam_2_average_저온)
print ("유효성분 :", 시료명3_저온)
print ("표준품 순도(%) ", std_content3_저온, "표준품 무게 ",std_g3_저온,"시료 무게 ",sam_3_1_g_저온," " ,sam_3_2_g_저온," ",sam_3_3_g_저온)
print('저온 안정성 시험 시료3', sam_3_1_content_저온, sam_3_2_content_저온, sam_3_3_content_저온, sam_3_average_저온)
print('저온 분해율:', 시료1경변분해율저온, 시료2경변분해율저온, 시료3경변분해율저온)

print ("품목의 물리성")
print(검사항목1_경변시작)
print(검사항목2_경변시작)
print(검사항목3_경변시작)
print(검사항목4_경변시작)
print(검사항목5_경변시작)

print ("4. 시험결과 요약")
print ("가. 가열안정성 시험에서 유효성분의 경시분해율은 ", 시료명1_경변3차, 시료1경변분해율3차, 시료명2_경변3차, 시료2경변분해율3차, 시료명3_경변3차, 시료3경변분해율3차,
       시료명1_경변4차, 시료1경변분해율4차, 시료명2_경변4차, 시료2경변분해율4차, 시료명3_경변4차, 시료3경변분해율4차, "이고")
print ("물리성은 양호하였다.")
print ("나. 이상의 결과를 종합하여 ",시험물질_경변시작,"의 제제안정성은 양호한 것으로 판단되며,")
print ("품목의 약효보증기간을 0년으로 설정하여도 타당할 것으로 판단되었다.")

print ("5. 첨부자료")
print ("가. 유효성분 분석 성적계산서 및 크로마토그램(가열안정성)")







