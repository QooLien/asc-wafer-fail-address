"""Remove unused worksheets from Artifact export (worksheet deletion API unavailable)."""
import sys
import zipfile
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

source, output = sys.argv[1:]
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
with zipfile.ZipFile(source) as z:
    parts = {n: z.read(n) for n in z.namelist()}
book = ET.fromstring(parts['xl/workbook.xml'])
rels = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
types = ET.fromstring(parts['[Content_Types].xml'])
sheets = book.find(f'{{{S}}}sheets')
for sheet in list(sheets):
    if sheet.get('name') not in ('Instructions', 'Import Log'):
        continue
    rid = sheet.get(f'{{{R}}}id')
    rel = next(r for r in rels if r.get('Id') == rid)
    target = rel.get('Target')
    part = target.lstrip('/') if target.startswith('/') else str(PurePosixPath('xl') / target)
    parts.pop(part)
    rel_part = str(PurePosixPath(part).parent / '_rels' / (PurePosixPath(part).name + '.rels'))
    assert rel_part not in parts, 'Unexpected related objects on removed worksheet'
    for item in list(types):
        if item.get('PartName') == '/' + part:
            types.remove(item)
    sheets.remove(sheet)
    rels.remove(rel)
next(s for s in sheets if s.get('name') == 'Fail Addresses').set('name', 'Fail Address')
assert [s.get('name') for s in sheets] == ['Config', 'Fail Address']
# Add unsupported vertical text and standalone AutoFilter using standard OpenXML.
from copy import deepcopy
styles = ET.fromstring(parts['xl/styles.xml'])
xfs = styles.find(f'{{{S}}}cellXfs')
page = ET.fromstring(parts['xl/worksheets/sheet2.xml'])
cell = next(c for c in page.iter(f'{{{S}}}c') if c.get('r') == 'G1')
style = deepcopy(xfs[int(cell.get('s','0'))])
alignment = style.find(f'{{{S}}}alignment')
if alignment is None:
    alignment = ET.SubElement(style,f'{{{S}}}alignment')
alignment.set('textRotation','90')
style.set('applyAlignment','1')
cell.set('s',str(len(xfs)))
xfs.append(style)
xfs.set('count',str(len(xfs)))
old = page.find(f'{{{S}}}autoFilter')
if old is not None:
    page.remove(old)
# autoFilter follows sheetData / sheetCalcPr / sheetProtection / protectedRanges / scenarios.
flt = ET.Element(f'{{{S}}}autoFilter',ref='A1:G1')
pos = list(page).index(page.find(f'{{{S}}}sheetData')) + 1
while pos < len(page) and page[pos].tag.rsplit('}',1)[-1] in ('sheetCalcPr','sheetProtection','protectedRanges','scenarios'):
    pos += 1
page.insert(pos,flt)
ET.register_namespace('',S)
parts['xl/styles.xml'] = ET.tostring(styles,encoding='utf-8',xml_declaration=True)
parts['xl/worksheets/sheet2.xml'] = ET.tostring(page,encoding='utf-8',xml_declaration=True)
for name, root, ns in [('xl/workbook.xml',book,S), ('xl/_rels/workbook.xml.rels',rels,P), ('[Content_Types].xml',types,C)]:
    ET.register_namespace('', ns)
    ET.register_namespace('r', R)
    parts[name] = ET.tostring(root,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
    for name, value in parts.items():
        z.writestr(name,value)
print('PASS: removed Instructions and Import Log; retained Config and Fail Address; vertical TNO and AutoFilter.')
