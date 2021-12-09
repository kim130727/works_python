
#replace file ext.
import glob
import os.path
import win32com.client

files = glob.glob('c:\\data automation\\COA\\*.xls')
s = os.path.split(files[0])
print (s[0])

for x in files:
    if not os.path.isdir(x):
        print (x.rstrip(".xls"))
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(x)
        ws = wb.ActiveSheet
        print(ws.Cells(22, 3).Value)
        if ws.Cells(22, 3).Value == 'Dongbu Hitek Co., Ltd.':
            print ('동부하이텍')
            cell = ws.Cells(1, 2)
            pic = ws.Pictures().Insert(r"c:\data automation\Dongbuhitek.jpg")
            pic.Left = cell.Left
            pic.Top = cell.Top

        elif ws.Cells(22, 3).Value == 'Dongbu Farm Hannong Co., Ltd.':
            print ('동부팜한농')
            cell = ws.Cells(1, 2)
            pic = ws.Pictures().Insert(r"c:\data automation\Dongbufarmhannong.jpg")
            pic.Left = cell.Left
            pic.Top = cell.Top

        elif ws.Cells(22, 3).Value == 'Dongbu Hannong Co., Ltd.':
            print ('동부한농')
            cell = ws.Cells(1, 2)
            pic = ws.Pictures().Insert(r"c:\data automation\Dongbuhannong.jpg")
            pic.Left = cell.Left
            pic.Top = cell.Top

        else:
            print('오류')

        wb.SaveAs(x.rstrip(".xls") + '.xlsx', FileFormat = 51)
        # FileFormat = 51 is for .xlsx extension
        # FileFormat = 56 is for .xls extension
        excel.Quit()