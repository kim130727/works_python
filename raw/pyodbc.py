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

nm = input('시험번호를 입력해 주세요 :  ')
mn = input('번호를 입력해 주세요. 1. 이화학, 2. P1, 3. P2?  :  ')

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


