"""Add supported OpenXML controls missing in Artifact's worksheet formatting API."""
from copy import deepcopy
from xml.etree import ElementTree as ET
import zipfile,sys
source,output=sys.argv[1:]
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
with zipfile.ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
styles=ET.fromstring(parts['xl/styles.xml']);xfs=styles.find(f'{{{S}}}cellXfs')
page=ET.fromstring(parts['xl/worksheets/sheet2.xml'])
row=next(r for r in page.find(f'{{{S}}}sheetData') if r.get('r')=='1')
for cell in row:
    if cell.get('r') in ('A1','B1','C1','D1','E1','F1'):continue
    xf=deepcopy(xfs[int(cell.get('s','0'))]);alignment=xf.find(f'{{{S}}}alignment')
    if alignment is None:alignment=ET.SubElement(xf,f'{{{S}}}alignment')
    alignment.set('textRotation','90');alignment.set('wrapText','1');xf.set('applyAlignment','1')
    cell.set('s',str(len(xfs)));xfs.append(xf)
xfs.set('count',str(len(xfs)))
last=next(reversed(list(row))).get('r')
page.insert(list(page).index(page.find(f'{{{S}}}sheetData'))+1,ET.Element(f'{{{S}}}autoFilter',ref='A1:'+last))
ET.register_namespace('',S);ET.register_namespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
parts['xl/styles.xml']=ET.tostring(styles,encoding='utf-8',xml_declaration=True)
parts['xl/worksheets/sheet2.xml']=ET.tostring(page,encoding='utf-8',xml_declaration=True)
book=ET.fromstring(parts['xl/workbook.xml'])
if book.find(f'{{{S}}}bookViews') is None:
    views=ET.Element(f'{{{S}}}bookViews')
    ET.SubElement(views,f'{{{S}}}workbookView',activeTab='0')
    book.insert(list(book).index(book.find(f'{{{S}}}sheets')),views)
for sheet in book.find(f'{{{S}}}sheets'):
    if sheet.get('name')=='Import Index':sheet.set('state','veryHidden')
parts['xl/workbook.xml']=ET.tostring(book,encoding='utf-8',xml_declaration=True)
for name in list(parts):
    if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
        sheet=ET.fromstring(parts[name]);views=sheet.find(f'{{{S}}}sheetViews')
        if views is None:
            views=ET.Element(f'{{{S}}}sheetViews');pos=0
            while pos<len(sheet) and sheet[pos].tag in (f'{{{S}}}sheetPr',f'{{{S}}}dimension'):pos+=1
            sheet.insert(pos,views)
        if len(views)==0:ET.SubElement(views,f'{{{S}}}sheetView',workbookViewId='0')
        for view in views:view.set('showGridLines','0')
        parts[name]=ET.tostring(sheet,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
    for n,v in parts.items():z.writestr(n,v)
