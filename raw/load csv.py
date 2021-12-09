
import csv, sys

filename = '완파.csv'
f = open(filename, 'rb')
reader = csv.reader(f)
try:
    for row in reader:
        for r in row:
            print (r.decode('euckr').encode('utf-8'))

except csv.Error as e:
    sys.exit('file %s, line %d: %s' % (filename, reader.line_num, e))

f.close()