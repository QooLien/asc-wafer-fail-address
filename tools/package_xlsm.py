"""Attach a validated MS-OVBA payload to an Artifact Tool-authored workbook.

Usage: python package_xlsm.py base.xlsx vbaProject.bin output.xlsm
No spreadsheet values/styles are authored here; only OPC macro relationships,
content types and Excel host code names are added.
"""
import pathlib
import sys
import zipfile
from xml.etree import ElementTree as ET

base, binary, output = map(pathlib.Path, sys.argv[1:])
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('', NS)
with zipfile.ZipFile(base) as z:
    entries = {name: z.read(name) for name in z.namelist()}
ct = ET.fromstring(entries['[Content_Types].xml'])
for node in ct:
    if node.get('PartName') == '/xl/workbook.xml':
        node.set('ContentType', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
ET.SubElement(ct, f'{{{CT}}}Override', PartName='/xl/vbaProject.bin', ContentType='application/vnd.ms-office.vbaProject')
entries['[Content_Types].xml'] = ET.tostring(ct, encoding='utf-8', xml_declaration=True)
rels = ET.fromstring(entries['xl/_rels/workbook.xml.rels'])
ET.SubElement(rels, f'{{{REL}}}Relationship', Id='rIdVBAProject', Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject', Target='vbaProject.bin')
entries['xl/_rels/workbook.xml.rels'] = ET.tostring(rels, encoding='utf-8', xml_declaration=True)
book = ET.fromstring(entries['xl/workbook.xml'])
props = book.find(f'{{{NS}}}workbookPr')
if props is None:
    props = ET.Element(f'{{{NS}}}workbookPr')
    # workbookPr follows fileVersion and fileSharing, before all other elements.
    idx = 0
    while idx < len(book) and book[idx].tag.rsplit('}', 1)[-1] in ('fileVersion', 'fileSharing'):
        idx += 1
    book.insert(idx, props)
props.set('codeName', 'ThisWorkbook')
entries['xl/workbook.xml'] = ET.tostring(book, encoding='utf-8', xml_declaration=True)
sheet_count = len(book.find(f'{{{NS}}}sheets'))
for idx in range(1, sheet_count + 1):
    name = f'xl/worksheets/sheet{idx}.xml'
    sheet = ET.fromstring(entries[name])
    props = sheet.find(f'{{{NS}}}sheetPr')
    if props is None:
        props = ET.Element(f'{{{NS}}}sheetPr')
        sheet.insert(0, props)
    props.set('codeName', f'Sheet{idx}')
    entries[name] = ET.tostring(sheet, encoding='utf-8', xml_declaration=True)
entries['xl/vbaProject.bin'] = binary.read_bytes()
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, content in entries.items():
        z.writestr(name, content)
with zipfile.ZipFile(output) as z:
    assert z.testzip() is None
    for name in entries:
        if name.endswith('.xml') or name.endswith('.rels'):
            ET.fromstring(z.read(name))
print(f'Packaged XLSM: VBA relationship, macro content type, {sheet_count + 1} host code names, XML/ZIP integrity verified.')
