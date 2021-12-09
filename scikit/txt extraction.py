
import re

f = open('C:\data automation\\formulation.txt', 'rt')
lines = f.readlines()

ar = re.compile(" +\w+")

n = 0

while n < 10000:

    area = ar.findall(lines[n])

    print (str(area))

    n = n+1