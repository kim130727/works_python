
import requests
import os
import re
import sys

filename = "c:data automation\\difenoconazole.xml"

f = open(filename, 'r', encoding='UTF8')
lines = f.readlines()

parse = lines[0]
print (parse)
