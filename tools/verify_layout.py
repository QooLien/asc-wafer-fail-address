"""Verify saved formatting and the native button binding without opening Excel."""
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET

root = Path(__file__).resolve().parents[1]
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(root / 'templates/ASC_Fail_Address.xlsm') as archive:
    styles = ET.fromstring(archive.read('xl/styles.xml'))
    fonts = list(styles.find(S + 'fonts'))
    formats = list(styles.find(S + 'cellXfs'))
    count = 0
    for index in range(1, 5):
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
                assert float(row.get('ht', '15')) == 26, ('row height', row.get('r'))
    config = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    assert config.find(S + 'legacyDrawing') is not None
    vml = ET.fromstring(archive.read('xl/drawings/vmlDrawing1.vml'))
    X = '{urn:schemas-microsoft-com:office:excel}'
    assert vml.find('.//' + X + 'FmlaMacro').text == '[0]!ImportASCFailAddresses'
    assert vml.find('.//' + X + 'TextHAlign').text == 'Center'
    assert vml.find('.//' + X + 'TextVAlign').text == 'Center'
    assert vml.find('.//font').get('face') == 'Calibri'
print(f'PASS: {count} populated cells use centered Calibri; Fail Addresses rows=26 pt; Forms button centered and macro-linked.')
