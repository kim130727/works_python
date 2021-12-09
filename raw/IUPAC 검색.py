
import codecs
import win32com.client
import pubchempy as pcp
import os
import glob

n = 0

while n < 150:

    n = n + 1

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open('C:\data automation\coa.xlsx')
    ws = wb.ActiveSheet

    m = float(n)+2

    productname = ws.Cells(m,3).Value
    commonname = ws.Cells(m,4).Value
    relatedproduct = ws.Cells(m,5).Value
    lotnumber = ws.Cells(m,6).Value
    manufacturer = ws.Cells(m,7).Value
    purity = ws.Cells(m,8).Value
    사명 = ws.Cells(m,9).Value
    담당자 = ws.Cells(m,10).Value
    analysisdate = ws.Cells(m,11).Value
    expirationdate = ws.Cells(m,12).Value

    try:
        CID = pcp.get_compounds(relatedproduct, 'name')
        IUPAC = CID[0].iupac_name

        if IUPAC == None:
            IUPAC = ""

    except IndexError:
        IUPAC = ""

    analysisyear = analysisdate[0:4]
    analysismonth = analysisdate[6:7]
    analysisday = analysisdate[8:10]
    expirationyear = expirationdate[0:4]
    expirationday = expirationdate[8:10]

    print (analysismonth)

    if analysismonth == '1' :
        month = "Jan."
    elif analysismonth == '2' :
        month = "Feb."
    elif analysismonth == '3':
        month = "Mar."
    elif analysismonth == '4':
        month = "Apr."
    elif analysismonth == '5':
        month = "May"
    elif analysismonth == '6':
        month = "June"
    elif analysismonth == '7':
        month = "July"
    elif analysismonth == '8':
        month = "Aug."
    elif analysismonth == '9':
        month = "Sep."
    elif analysismonth == '10':
        month = "Oct."
    elif analysismonth == '11':
        month = "Nov."
    elif analysismonth == '12':
        month = "Dec."

    print (month)
    analysis = analysisyear + " " + month + " " + analysisday
    expiration = expirationyear + " " + month + " " + expirationday

    print (analysis)
    print (expiration)

    print (productname, " ", commonname, " ", relatedproduct, " ", lotnumber, " ", manufacturer, " ", IUPAC)
    print (purity, " ", 사명, " ", 담당자, " ", analysisdate, " ", analysisyear, " ", expirationdate)

    f = codecs.open("C:\data automation\COA\IUPAC.txt", 'a', 'utf-8')
    f.write(relatedproduct + "  " + IUPAC+'\n')

