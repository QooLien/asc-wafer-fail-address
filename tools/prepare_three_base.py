"""Add standard OpenXML header rotations / AutoFilter unsupported by Artifact API."""
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET
import sys,zipfile
source,output=sys.argv[1:]
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
with zipfile.ZipFile(source) as z: parts={n:z.read(n) for n in z.namelist()}
styles=ET.fromstring(parts['xl/styles.xml']);xfs=styles.find(f'{{{S}}}cellXfs')
page=ET.fromstring(parts['xl/worksheets/sheet2.xml'])
for cell in page.iter(f'{{{S}}}c'):
 if cell.get('r') not in ('G1','H1','I1'): continue
 style=deepcopy(xfs[int(cell.get('s','0'))]);a=style.find(f'{{{S}}}alignment')
 if a is None:a=ET.SubElement(style,f'{{{S}}}alignment')
 a.set('textRotation','90');a.set('wrapText','1');style.set('applyAlignment','1')
 cell.set('s',str(len(xfs)));xfs.append(style)
xfs.set('count',str(len(xfs)))
flt=ET.Element(f'{{{S}}}autoFilter',ref='A1:I1')
page.insert(list(page).index(page.find(f'{{{S}}}sheetData'))+1,flt)
ET.register_namespace('',S);ET.register_namespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
parts['xl/styles.xml']=ET.tostring(styles,encoding='utf-8',xml_declaration=True)
parts['xl/worksheets/sheet2.xml']=ET.tostring(page,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
