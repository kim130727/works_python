# pdfx usage: http://pdfx.cs.man.ac.uk/usage
# requests docs: http://docs.python-requests.org/en/latest/user/quickstart/#post-a-multipart-encoded-file
import requests # get it from http://python-requests.org or do 'pip install requests'
import os
import re
import sys

url = "http://pdfx.cs.man.ac.uk"
filename = "C:\data automation\\2.부자재\부자재_코씰\GHS_MSDS\\NK-8020.pdf"
fin = open(filename, 'rb')
files = {'file': fin}

print ('Sending', filename, 'to', url)
r = requests.post(url, files=files, headers={'Content-Type':'application/pdf'})
print ('Got status code', r.status_code)

fout = open(filename.split('.')[0] + '.xml', 'w')
fout.write(str(r.content))

fout.close()
print ('Written to', filename.split('.')[0] + '.xml')

f = open(filename.split('.')[0] + '.xml', 'rt')
lines = f.readlines()

print (lines)
parse = lines[0]
parse = re.sub('<',' ', parse)
parse = re.sub('>',' ', parse)
parse = re.sub('"',' ', parse)
parse = re.sub('marker type',' ', parse)
parse = re.sub('region class',' ', parse)
parse = re.sub('DoCO:TextChunk',' ', parse)
parse = re.sub('confidence',' ', parse)
parse = re.sub('possible',' ', parse)
parse = re.sub('outsider class',' ', parse)
parse = re.sub('block',' ', parse)
parse = re.sub('region',' ', parse)
parse = re.sub('DoCO:TextBox',' ', parse)
parse = re.sub('article',' ', parse)
parse = re.sub('DoCO:FigureBox',' ', parse)
parse = re.sub('DoCO:Title',' ', parse)
parse = re.sub('DoCO:FrontMatter',' ', parse)
parse = re.sub('/',' ', parse)
parse = re.sub('&gt;','>', parse)
parse = re.sub('type=','', parse)
parse = re.sub('sidenote','', parse)
parse = re.sub('outsider','', parse)
parse = re.sub('id=','', parse)
parse = re.sub('header','', parse)
parse = re.sub('DoCO:Figure','', parse)
parse = re.sub('front class','', parse)
parse = re.sub('title class','', parse)
parse = re.sub('footer','', parse)
parse = re.sub('=','', parse)
parse = re.sub('DoCO:BodyMatter','', parse)
parse = re.sub('                 ',' ', parse)
parse = re.sub('                ',' ', parse)
parse = re.sub('               ',' ', parse)
parse = re.sub('              ',' ', parse)
parse = re.sub('             ',' ', parse)
parse = re.sub('            ',' ', parse)
parse = re.sub('           ',' ', parse)
parse = re.sub('          ',' ', parse)
parse = re.sub('         ',' ', parse)
parse = re.sub('        ',' ', parse)
parse = re.sub('       ',' ', parse)
parse = re.sub('      ',' ', parse)
parse = re.sub('     ',' ', parse)
parse = re.sub('    ',' ', parse)
parse = re.sub('   ',' ', parse)
parse = re.sub('  ',' ', parse)
print (parse)

area = re.compile("Area \d+[.]\d\d\d")
rt = re.compile("rt \d+[.]\d\d")
vial = re.compile("\d+[:]\w+[,]\d+")
name = re.compile("\w+[_]\w+[_]\w+[_]\w+")
time1 = re.compile("Last Altered[:] \w+[,] \w+ \w+[,] \d+ \d+[:]\d+[:]\d+")
time2 = re.compile("Printed[:] \w+[,] \w+ \w+[,] \d+ \d+[:]\d+[:]\d+")
method = re.compile("Method[:] \w+[.mdb]")
page = re.compile("page \d+")
series = re.compile("Name: \S+[,] Date: \S+[,] Time: \S+[,] Vial: \S+")
operator = re.compile("Operator : \S+ \S+")


area = area.findall(parse)
rt = rt.findall(parse)
vial = vial.findall(parse)
name = name.findall(parse)
time1 = time1.findall(parse)
time2 = time2.findall(parse)
method = method.findall(parse)
page = page.findall(parse)
series = series.findall(parse)
operator = operator.findall(parse)
rt = rt[1:]


print (area)
print (rt)
print (vial)
print (name)
print (time1)
print (time2)
print (method)
print (page)
print (series)
print (operator)

n = 0

try:
 while n < 1000:
     f = open('c:\data automation\\report.txt', 'a')
     f.write(time1[n])
     f.write(" ")
     f.write(time2[n])
     f.write(" ")
     f.write(operator[n])
     f.write(" ")
     f.write(series[n])
     f.write(" ")
     f.write(area[n])
     f.write('\n')
     f.close()
     n = n+1
except:
    pass

