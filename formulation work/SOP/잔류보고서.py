
import requests
import os
import re
import sys

filename = "c:data automation\\difenoconazole.xml"

f = open(filename, 'r', encoding='UTF8')
lines = f.readlines()

n = 0

try:
    while n < 500000:
        parse = lines[n]
        parse = re.sub('디페노코나졸','종근이코나졸', parse)
        print ("checking")
        f2 = open('c:data automation\\final.xml', 'a', encoding='UTF8')
        f2.write(parse)
        f2.close()
        n = n+1
except:
    pass

f.close()
print (lines[0])

