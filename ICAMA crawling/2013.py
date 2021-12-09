import requests
import json
import ast
import datetime
import pandas as pd
import re

n =1

while n < 10000:

    f = open("C:/data automation/PD2013.txt",'a')

    url = "http://www.chinapesticide.org.cn/myquery/querydetail_en?pdno="+"PD2013"+'{0:04}'.format(n)+"&___t0.061850364002740355"

    source_code = requests.get(url)
    plain_text = source_code.text

    jsonString = json.dumps(plain_text)

    parse = re.sub(' ', '', jsonString)

    result = re.findall("Formulation..................................................................................................................................................................", parse)

    code = parse[4637:4647]
    formulation = str(result)[232:234]
    print (code, " ", formulation)

    f.write(code)
    f.write(" ")
    f.write(formulation)
    f.write("\n")
    n = n+1





