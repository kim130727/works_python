
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


    f = codecs.open("C:\data automation\COA\coa.xml", 'w', 'utf-8')

    f.write('<?xml version="1.0"?>\n')
    f.write('<?mso-application progid="Excel.Sheet"?>\n')
    f.write('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
    f.write(' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
    f.write(' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
    f.write(' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n')
    f.write(' xmlns:html="http://www.w3.org/TR/REC-html40">\n')
    f.write(' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n')
    f.write('  <Author>user</Author>\n')
    f.write('  <LastAuthor>USER</LastAuthor>\n')
    f.write('  <LastPrinted>2015-07-06T23:35:13Z</LastPrinted>\n')
    f.write('  <Created>2015-07-03T07:44:58Z</Created>\n')
    f.write('  <LastSaved>2016-11-04T08:24:56Z</LastSaved>\n')
    f.write('  <Version>14.00</Version>\n')
    f.write(' </DocumentProperties>\n')
    f.write(' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">\n')
    f.write('  <AllowPNG/>\n')
    f.write(' </OfficeDocumentSettings>\n')
    f.write(' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">\n')
    f.write('  <WindowHeight>9030</WindowHeight>\n')
    f.write('  <WindowWidth>11490</WindowWidth>\n')
    f.write('  <WindowTopX>195</WindowTopX>\n')
    f.write('  <WindowTopY>60</WindowTopY>\n')
    f.write('  <ProtectStructure>False</ProtectStructure>\n')
    f.write('  <ProtectWindows>False</ProtectWindows>\n')
    f.write(' </ExcelWorkbook>\n')
    f.write(' <Styles>\n')
    f.write('  <Style ss:ID="Default" ss:Name="Normal">\n')
    f.write('   <Alignment ss:Vertical="Center"/>\n')
    f.write('   <Borders/>\n')
    f.write('   <Font ss:FontName="맑은 고딕" x:CharSet="129" x:Family="Modern" ss:Size="11"\n')
    f.write('    ss:Color="#000000"/>\n')
    f.write('   <Interior/>\n')
    f.write('   <NumberFormat/>\n')
    f.write('   <Protection/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590656">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590676">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590696">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590716">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('   <NumberFormat ss:Format="Percent"/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590776">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590836">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590856">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590876">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('   <NumberFormat/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590896">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('   <NumberFormat/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590916">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="m61590936">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s16">\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s18">\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s19">\n')
    f.write('   <Alignment ss:Vertical="Bottom"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s20">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s21">\n')
    f.write('   <Font ss:FontName="맑은 고딕" x:CharSet="129" x:Family="Modern" ss:Size="11"\n')
    f.write('    ss:Color="#000000"/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s22">\n')
    f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s47">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n')
    f.write('   <Borders>\n')
    f.write('    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n')
    f.write('   </Borders>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s53">\n')
    f.write('   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>\n')
    f.write('   <Font ss:FontName="맑은 고딕" x:CharSet="129" x:Family="Modern" ss:Size="8"\n')
    f.write('    ss:Color="#000000"/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s54">\n')
    f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
    f.write('   <Font ss:FontName="맑은 고딕" x:CharSet="129" x:Family="Modern" ss:Size="24"\n')
    f.write('    ss:Color="#000000"/>\n')
    f.write('  </Style>\n')
    f.write('  <Style ss:ID="s55">\n')
    f.write('   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n')
    f.write('   <Font ss:FontName="맑은 고딕" x:CharSet="129" x:Family="Modern" ss:Size="11"\n')
    f.write('    ss:Color="#000000"/>\n')
    f.write('  </Style>\n')
    f.write(' </Styles>\n')
    f.write(' <Worksheet ss:Name="Sheet1">\n')
    f.write('  <Names>\n')
    f.write('   <NamedRange ss:Name="Print_Area" ss:RefersTo="=Sheet1!R1C2:R30C5"/>\n')
    f.write('  </Names>\n')
    f.write('  <Table ss:ExpandedColumnCount="5" ss:ExpandedRowCount="29" x:FullColumns="1"\n')
    f.write('   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="16.5">\n')
    f.write('   <Column ss:AutoFitWidth="0" ss:Width="20.25"/>\n')
    f.write('   <Column ss:AutoFitWidth="0" ss:Width="154.5" ss:Span="3"/>\n')
    f.write('   <Row ss:Index="4" ss:AutoFitHeight="0" ss:Height="36">\n')
    f.write('    <Cell ss:Index="2" ss:MergeAcross="3" ss:StyleID="s53"><Data ss:Type="String">')
    f.write(사명)
    f.write(' Co., Ltd., ')

    if analysisyear == "2015":
        f.write('Crop Protection R&D Center, &#10;39-23, Dongan-ro 1113beon-gil, Yeonmu-eup, Nonsan-si, Chungcheongnam-do, Korea &#10;Tel : +82 41 730 9125,   Fax : +82 41 742 8510&#10;</Data><NamedCell\n')

    else :
        f.write('Dongbu Advanced Research Institute, &#10;229 munji-ro, Yuseong-gu, Daejeon 305-708, Korea &#10;Tel : +82 42 866 8114,   Fax : +82 42 861 1583&#10;</Data><NamedCell\n')

    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="48.5625"/>\n')
    f.write('   <Row ss:Height="38.25">\n')
    f.write('    <Cell ss:Index="2" ss:MergeAcross="3" ss:StyleID="s54"><Data ss:Type="String">Certificate of Analysis</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row>\n')
    f.write('    <Cell ss:Index="2" ss:MergeAcross="3" ss:StyleID="s55"><Data ss:Type="String">')
    f.write('Lot Number : ')
    f.write(lotnumber)
    f.write('</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="31.875"/>\n')
    f.write('   <Row>\n')
    f.write('    <Cell ss:Index="2"><Data ss:Type="String"> General &amp; Analytical Information</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Product Name</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s16"><Data ss:Type="String">')
    f.write(productname)
    f.write('</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s16"><Data ss:Type="String"> Synonym</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s16"><Data ss:Type="String" x:Ticked="1">-</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Common Name</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s16"><Data ss:Type="String">')
    f.write(commonname)
    f.write(' </Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s16"><Data ss:Type="String"> Related Product</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s16"><Data ss:Type="String">')
    f.write(relatedproduct)
    f.write(' </Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Lot Number</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590856"><Data ss:Type="String">')
    f.write(lotnumber)
    f.write('</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Manufaturer</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590776"><Data ss:Type="String">')
    f.write(manufacturer)
    f.write('</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:MergeDown="1" ss:StyleID="s47"><Data ss:Type="String"> Chemical Name(IUPAC)</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="m61590936"><Data\n')
    f.write('      ss:Type="String">')
    f.write(str(IUPAC))
    f.write('  </Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125"/>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Purity</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590716"><Data ss:Type="String">')
    f.write(str(purity))
    f.write('%</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Analysis Date</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590876"><Data ss:Type="String"\n')
    f.write('      x:Ticked="1">')
    f.write(analysis)
    f.write('</Data><NamedCell ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:MergeDown="1" ss:StyleID="m61590916"><Data\n')
    f.write('      ss:Type="String"> Expiration Date</Data><NamedCell ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590896"><Data ss:Type="String"\n')
    f.write('      x:Ticked="1">')
    f.write(expiration)
    f.write('</Data><NamedCell ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="53.4375">\n')
    f.write('    <Cell ss:Index="3" ss:MergeAcross="2" ss:StyleID="m61590836"><Data\n')
    f.write('      ss:Type="String">It has been established that this chemical sustance was stable for 5 years under the described storage conditions, referenced by CRP-CHM-EXP-002-00. This sustance would be stable for the same term the date of analysis for the purity determination.</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Analytical Method</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590656"><Data ss:Type="String">HPLC</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Storage Conditions</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590676"><Data ss:Type="String">Below 5°C, Dark condition</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="25.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s16"><Data ss:Type="String"> Testing Facility</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:MergeAcross="2" ss:StyleID="m61590696"><Data ss:Type="String">')
    f.write(사명)
    f.write(' Co., Ltd.</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="53.4375"/>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="150"/>\n')
    f.write('   <Row ss:AutoFitHeight="0" ss:Height="37.125">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s19"><Data ss:Type="String">(Signature)</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s18"><NamedCell ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:Index="5" ss:StyleID="s22"><Data ss:Type="String" x:Ticked="1">')
    f.write(analysis)
    f.write('</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row>\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s21"><Data ss:Type="String">Signed : ')
    f.write(담당자)
    f.write(' / Research Scientist</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:Index="5"><Data ss:Type="String">Issue date</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row ss:Index="28" ss:AutoFitHeight="0" ss:Height="39">\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s19"><Data ss:Type="String">(Signature)</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:StyleID="s18"><NamedCell ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:Index="5" ss:StyleID="s22"><Data ss:Type="String">')
    f.write(analysis)
    f.write('</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('   <Row>\n')
    f.write('    <Cell ss:Index="2" ss:StyleID="s20"><Data ss:Type="String">Approved : J. K. Eom / Senior Research Scientist</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('    <Cell ss:Index="5"><Data ss:Type="String">Issue date</Data><NamedCell\n')
    f.write('      ss:Name="Print_Area"/></Cell>\n')
    f.write('   </Row>\n')
    f.write('  </Table>\n')
    f.write('  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n')
    f.write('   <PageSetup>\n')
    f.write('    <Header x:Margin="0.31496062992125984"/>\n')
    f.write('    <Footer x:Margin="0.31496062992125984"/>\n')
    f.write('    <PageMargins x:Bottom="0.74803149606299213" x:Left="0.70866141732283472"\n')
    f.write('     x:Right="0.70866141732283472" x:Top="0.74803149606299213"/>\n')
    f.write('   </PageSetup>\n')
    f.write('   <FitToPage/>\n')
    f.write('   <Print>\n')
    f.write('    <ValidPrinterInfo/>\n')
    f.write('    <PaperSizeIndex>9</PaperSizeIndex>\n')
    f.write('    <Scale>78</Scale>\n')
    f.write('    <HorizontalResolution>600</HorizontalResolution>\n')
    f.write('    <VerticalResolution>600</VerticalResolution>\n')
    f.write('   </Print>\n')
    f.write('   <Zoom>90</Zoom>\n')
    f.write('   <Selected/>\n')
    f.write('   <Panes>\n')
    f.write('    <Pane>\n')
    f.write('     <Number>3</Number>\n')
    f.write('     <ActiveRow>4</ActiveRow>\n')
    f.write('     <ActiveCol>4</ActiveCol>\n')
    f.write('    </Pane>\n')
    f.write('   </Panes>\n')
    f.write('   <ProtectObjects>False</ProtectObjects>\n')
    f.write('   <ProtectScenarios>False</ProtectScenarios>\n')
    f.write('  </WorksheetOptions>\n')
    f.write(' </Worksheet>\n')
    f.write(' <Worksheet ss:Name="Sheet2">\n')
    f.write('  <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"\n')
    f.write('   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="16.5">\n')
    f.write('  </Table>\n')
    f.write('  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n')
    f.write('   <PageSetup>\n')
    f.write('    <Header x:Margin="0.3"/>\n')
    f.write('    <Footer x:Margin="0.3"/>\n')
    f.write('    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>\n')
    f.write('   </PageSetup>\n')
    f.write('   <ProtectObjects>False</ProtectObjects>\n')
    f.write('   <ProtectScenarios>False</ProtectScenarios>\n')
    f.write('  </WorksheetOptions>\n')
    f.write(' </Worksheet>\n')
    f.write(' <Worksheet ss:Name="Sheet3">\n')
    f.write('  <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"\n')
    f.write('   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="16.5">\n')
    f.write('  </Table>\n')
    f.write('  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n')
    f.write('   <PageSetup>\n')
    f.write('    <Header x:Margin="0.3"/>\n')
    f.write('    <Footer x:Margin="0.3"/>\n')
    f.write('    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>\n')
    f.write('   </PageSetup>\n')
    f.write('   <ProtectObjects>False</ProtectObjects>\n')
    f.write('   <ProtectScenarios>False</ProtectScenarios>\n')
    f.write('  </WorksheetOptions>\n')
    f.write(' </Worksheet>\n')
    f.write('</Workbook>\n')

    f.close()

    excel.Quit()

    fpath = "C:\data automation\COA\coa.xml"
    fpath_r = fpath.replace('coa', commonname+'_coa')
    os.rename(fpath, fpath_r)