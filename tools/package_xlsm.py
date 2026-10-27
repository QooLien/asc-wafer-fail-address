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
def set_content_type(part, content_type):
    node = next((item for item in ct if item.tag == f'{{{CT}}}Override' and item.get('PartName') == part), None)
    if node is None:
        node = ET.SubElement(ct, f'{{{CT}}}Override', PartName=part)
    node.set('ContentType', content_type)

# Artifact Tool may declare workbook.xml through Default Extension="xml",
# rather than an explicit workbook Override. Always add the explicit XLSM type.
set_content_type('/xl/workbook.xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
set_content_type('/xl/vbaProject.bin', 'application/vnd.ms-office.vbaProject')
ET.register_namespace('', CT)
entries['[Content_Types].xml'] = ET.tostring(ct, encoding='utf-8', xml_declaration=True)
rels = ET.fromstring(entries['xl/_rels/workbook.xml.rels'])
macro_rel = next((item for item in rels if item.get('Type') == 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'), None)
if macro_rel is None:
    macro_rel = ET.SubElement(rels, f'{{{REL}}}Relationship', Id='rIdVBAProject', Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject')
macro_rel.set('Target', 'vbaProject.bin')
ET.register_namespace('', REL)
entries['xl/_rels/workbook.xml.rels'] = ET.tostring(rels, encoding='utf-8', xml_declaration=True)
ET.register_namespace('', NS)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
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
    saved_ct = ET.fromstring(z.read('[Content_Types].xml'))
    overrides = {item.get('PartName'): item.get('ContentType') for item in saved_ct if item.tag == f'{{{CT}}}Override'}
    assert overrides['/xl/workbook.xml'] == 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
    assert overrides['/xl/vbaProject.bin'] == 'application/vnd.ms-office.vbaProject'
    assert z.read('xl/vbaProject.bin')[:8] == bytes.fromhex('d0cf11e0a1b11ae1')
print(f'Packaged XLSM: explicit macro workbook content type, VBA relationship/container, {sheet_count + 1} host code names, XML/ZIP integrity verified.')
