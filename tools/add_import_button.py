"""Add an Excel 2016 Forms button; sheet contents remain Artifact Tool-authored.

XlsxWriter is used only for the missing native Forms-control capability. A tiny
in-memory donor workbook supplies the VML control, not spreadsheet cells/styles.
Usage: python add_import_button.py styled-base.xlsx output-base.xlsx
"""
import io
import pathlib
import sys
import zipfile
from xml.etree import ElementTree as ET
import xlsxwriter

base, output = map(pathlib.Path, sys.argv[1:])
memory = io.BytesIO()
with xlsxwriter.Workbook(memory, {'in_memory': True}) as donor:
    sheet = donor.add_worksheet('Config')
    sheet.set_column('A:A', 16)
    sheet.set_column('B:B', 4)
    sheet.set_column('C:C', 26)
    sheet.set_column('D:D', 65)
    sheet.set_default_row(26)
    for row in range(1, 6): sheet.set_row(row, 32)
    for row in range(7, 10): sheet.set_row(row, 28)
    sheet.insert_button('C8', {'macro': 'ImportASCFailAddresses', 'caption': '選擇 ASC 並匯入', 'width': 600, 'height': 52})
with zipfile.ZipFile(io.BytesIO(memory.getvalue())) as donor_zip:
    vml_name = next(name for name in donor_zip.namelist() if name.endswith('.vml'))
    vml = donor_zip.read(vml_name).decode('utf-8')
    vml = vml.replace('face="Tahoma"', 'face="Calibri"').replace('size="160"', 'size="220"')
with zipfile.ZipFile(base) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
entries[vml_name] = vml.encode('utf-8')
rels_name = 'xl/worksheets/_rels/sheet1.xml.rels'
rels = ET.fromstring(entries[rels_name]) if rels_name in entries else ET.Element(f'{{{REL}}}Relationships')
ET.SubElement(rels, f'{{{REL}}}Relationship', Id='rIdImportButton', Type=R + '/vmlDrawing', Target='../drawings/' + pathlib.PurePosixPath(vml_name).name)
ET.register_namespace('', REL)
entries[rels_name] = ET.tostring(rels, encoding='utf-8', xml_declaration=True)
sheet = ET.fromstring(entries['xl/worksheets/sheet1.xml'])
legacy = ET.Element(f'{{{NS}}}legacyDrawing', {f'{{{R}}}id': 'rIdImportButton'})
ext = sheet.find(f'{{{NS}}}extLst')
if ext is None: sheet.append(legacy)
else: sheet.insert(list(sheet).index(ext), legacy)
ET.register_namespace('', NS)
ET.register_namespace('r', R)
entries['xl/worksheets/sheet1.xml'] = ET.tostring(sheet, encoding='utf-8', xml_declaration=True)
types = ET.fromstring(entries['[Content_Types].xml'])
ET.SubElement(types, f'{{{CT}}}Override', PartName='/' + vml_name, ContentType='application/vnd.openxmlformats-officedocument.vmlDrawing')
ET.register_namespace('', CT)
entries['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
    for name, content in entries.items(): archive.writestr(name, content)
assert '[0]!ImportASCFailAddresses' in vml
assert 'face="Calibri"' in vml
print('Added native Forms button: Config!C8, Calibri, centered caption, linked to ImportASCFailAddresses.')
