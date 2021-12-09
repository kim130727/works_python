
import requests
import os
import re
import sys

filename = "c:data automation\\18-RC020_B_F.txt"
fin = open(filename, 'rb')

lines = fin.read().splitlines()
fin.close()

parse = str(lines)
print (1)
print (parse)

parse = re.sub('b''','', parse)
parse = re.sub('[=+#/\?^$@*\"※~&%ㆍ!』\\‘|\(\)\[\]\<\>`\'…》]', '', parse)
print (parse)

operator = re.compile("Operator : \S+ \S+")
area = re.compile("Area , \d+.\d+")
name = re.compile("Name: \S+")
date = re.compile("Date: \S+")
time = re.compile("Time: \S+")
vial = re.compile("Vial: \S+")
rt = re.compile("RT , \d+.\d+")

operator = operator.findall(parse)
area = area.findall(parse)
area = area[2:]
name = name.findall(parse)
date = date.findall(parse)
time = time.findall(parse)
vial = vial.findall(parse)
rt = rt.findall(parse)
rt = rt[1:]

print (operator)
print (name)
print (date)
print (time)
print (vial)
print (rt)
print (area)

n = 0

try:
 while n < 1000:
     f = open('c:\data automation\\report.txt', 'a')
     f.write(operator[n])
     f.write(" ")
     f.write(name[n])
     f.write(" ")
     f.write(date[n])
     f.write(" ")
     f.write(time[n])
     f.write(" ")
     f.write(vial[n])
     f.write(" ")
     f.write(rt[n])
     f.write(" ")
     f.write(area[n])
     f.write('\n')
     f.close()
     n = n+1
except:
    pass
