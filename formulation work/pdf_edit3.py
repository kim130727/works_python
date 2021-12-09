# pdfx usage: http://pdfx.cs.man.ac.uk/usage
# requests docs: http://docs.python-requests.org/en/latest/user/quickstart/#post-a-multipart-encoded-file
import requests # get it from http://python-requests.org or do 'pip install requests'
import os
import re
import sys

direct = "C:\data automation"

def search(dirname):
    flist = os.listdir(dirname)
    for f in flist:
        next = os.path.join(dirname, f)
        print ("ha", next)
        if os.path.isdir(next):
            print("search")
            search(next)
        else:
            print("dofilefork")
            doFileWork(next)

def doFileWork(filename):
    ext = os.path.splitext(filename)[-1]
    print (ext)
    if ext == '.pdf':

        url = "http://pdfx.cs.man.ac.uk"

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

        f = open('c:\data automation\msds.txt', 'a')

        f.write(lines)

        f.close()

        os.remove(filename.split('.')[0] + '.xml')

    else:
        sys.exit()

print (search(direct))