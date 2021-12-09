
import PyPDF2
pdf_file = open('C:\data automation\\2.pdf', 'rb')
read_pdf = PyPDF2.PdfFileReader(pdf_file)
number_of_pages = read_pdf.getNumPages()
page = read_pdf.getPage(0)
page_content = page.extractText()
print (page_content.encode('utf-8'))
print (type(page_content.encode('utf-8')))

f = open('c:\data automation\msds.txt', 'a')
f.write(str(page_content.encode('utf-8')))
f.close()