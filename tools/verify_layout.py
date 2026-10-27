"""Verify saved formatting and the native button binding without opening Excel."""
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET

root = Path(__file__).resolve().parents[1]
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
workbook_path = Path(sys.argv[1]) if len(sys.argv) > 1 else root / 'templates/ASC_Fail_Address_Fast.xlsm'
with zipfile.ZipFile(workbook_path) as archive:
    book = ET.fromstring(archive.read('xl/workbook.xml'))
    assert [node.get('name') for node in book.find(S + 'sheets')] == ['Config', 'Fail Address']
    cfg = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    for cell in cfg.iter(S + 'c'):
        if cell.get('r') in {f'{col}{row}' for col in 'CD' for row in range(11,17)}:
            assert cell.find(S+'v') is None and cell.find(S+'is') is None, cell.get('r')
    styles = ET.fromstring(archive.read('xl/styles.xml'))
    fonts = list(styles.find(S + 'fonts'))
    formats = list(styles.find(S + 'cellXfs'))
    count = 0
    for index in range(1, 3):
        sheet = ET.fromstring(archive.read(f'xl/worksheets/sheet{index}.xml'))
        for cell in sheet.iter(S + 'c'):
            if cell.find(S + 'v') is None and cell.find(S + 'is') is None:
                continue
            fmt = formats[int(cell.get('s', '0'))]
            font = fonts[int(fmt.get('fontId', '0'))]
            assert font.find(S + 'name').get('val') == 'Calibri', (index, cell.get('r'), 'font')
            alignment = fmt.find(S + 'alignment')
            assert alignment is not None and alignment.get('horizontal') == 'center' and alignment.get('vertical') == 'center', (index, cell.get('r'), 'alignment')
            count += 1
        if index == 2:
            for row in sheet.find(S + 'sheetData'):
                assert float(row.get('ht', '15')) == (60 if row.get('r') == '1' else 26), ('row height', row.get('r'))
    output = ET.fromstring(archive.read('xl/worksheets/sheet2.xml'))
    assert output.find(S+'autoFilter').get('ref') == 'A1:G1'
    tno = next(c for c in output.iter(S+'c') if c.get('r')=='G1')
    assert formats[int(tno.get('s'))].find(S+'alignment').get('textRotation') == '90'
    config = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    assert config.find(S + 'legacyDrawing') is not None
    vml = ET.fromstring(archive.read('xl/drawings/vmlDrawing1.vml'))
    X = '{urn:schemas-microsoft-com:office:excel}'
    assert vml.find('.//' + X + 'FmlaMacro').text == '[0]!ImportASCFailAddresses'
    assert vml.find('.//' + X + 'TextHAlign').text == 'Center'
    assert vml.find('.//' + X + 'TextVAlign').text == 'Center'
    assert vml.find('.//font').get('face') == 'Calibri'
print(f'PASS: {count} populated cells use centered Calibri; Fail Address body rows=26 pt, header=60 pt, TNO rotation=90, filter enabled; Forms button centered and macro-linked.')
