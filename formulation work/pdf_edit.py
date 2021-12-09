# pdfx usage: http://pdfx.cs.man.ac.uk/usage
# requests docs: http://docs.python-requests.org/en/latest/user/quickstart/#post-a-multipart-encoded-file
import requests # get it from http://python-requests.org or do 'pip install requests'
import os
import re
import sys

n = input('읽어들일 총 물질의 수는? (1~4): ')
direct = input('검색할 디렉토리는? ')

def search(dirname):
    flist = os.listdir(dirname)
    for f in flist:
        next = os.path.join(dirname, f)
        if os.path.isdir(next):
            search(next)
        else:
            doFileWork(next)

def doFileWork(filename):
    ext = os.path.splitext(filename)[-1]
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

        id = re.compile("Injection Date.....................")
        it = re.compile("Injection Time..............")
        vial = re.compile("Vial \d+")
        ar = re.compile("\d+[.]\d+")

        Injectiondate = id.findall(lines[0])
        Injectiontime = it.findall(lines[0])
        vialno = vial.findall(lines[0])
        area = ar.findall(lines[0])

        print(lines)
        print(Injectiondate)
        print(Injectiontime)
        print(vialno)

        if n == '1':
            totalarea = area[-2] + ""

        elif n == '2':
            totalarea = area[-6] + " " + area[-2]

        elif n == '3':
            totalarea = area[-10] + " " + area[-6] + " " + area[-2]

        else:
            totalarea = area[-14] + " " + area[-10] + " " + area[-6] + " " + area[-2]

        print('area : ' + totalarea)

        f = open('c:\data automation\chromatogram_list.txt', 'a')

        f.write(filename)
        f.write(' ')
        f.write(Injectiondate[0])
        f.write(' ')
        f.write(Injectiontime[0])
        f.write(' ')
        f.write(vialno[0])
        f.write(' ')
        f.write("area " + totalarea + '\n')
        f.close()

        os.remove(filename.split('.')[0] + '.xml')

    else:
        sys.exit()

search(direct)