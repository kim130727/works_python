#-*- coding: utf-8 -*-

import requests # get it from http://python-requests.org or do 'pip install requests'
import os
import re
import sys
import PyPDF2

def search(dirname):
    filenames = os.listdir(dirname)
    for filename in filenames:
        full_filename = os.path.join(dirname, filename)
        ext = os.path.splitext(full_filename)[-1]
        if ext == '.pdf':
            pdf_file = open(full_filename, 'rb')
            read_pdf = PyPDF2.PdfFileReader(pdf_file)
            number_of_pages = read_pdf.getNumPages()
            n = 0
            try:
                while n < 100:
                    page = read_pdf.getPage(n)
                    page_content = page.extractText()

                    print (page_content)

                    print(page_content.encode('utf-8'))
                    print(type(page_content.encode('utf-8')))

                    msds = str(page_content.encode('utf-8'))
                    msds = re.sub('\n', '', msds)

                    f = open('c:\data automation\msds.txt', 'a')
                    f.write(msds)
                    f.write('\n')
                    f.close()
                    n = n+1
            except:
                pass
        else:
            pass

search("C:\data automation\\2.부자재\부자재_코씰\GHS_MSDS")