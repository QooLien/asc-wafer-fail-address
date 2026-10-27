"""Independent VBA extraction check. Requires oletools (development only)."""
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET
from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
def verify_macro_package(path):
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
        types = ET.fromstring(archive.read('[Content_Types].xml'))
        # Office/Windows OPC readers require the conventional unprefixed roots.
        assert b'<Types ' in archive.read('[Content_Types].xml'), 'Content-types root must use its default namespace'
        assert b'<Relationships ' in archive.read('xl/_rels/workbook.xml.rels'), 'Relationships root must use its default namespace'
        ns = '{http://schemas.openxmlformats.org/package/2006/content-types}'
        overrides = {node.get('PartName'): node.get('ContentType') for node in types if node.tag == ns + 'Override'}
        defaults = {node.get('Extension'): node.get('ContentType') for node in types if node.tag == ns + 'Default'}
        effective = overrides.get('/xl/workbook.xml', defaults.get('xml'))
        assert effective == 'application/vnd.ms-excel.sheet.macroEnabled.main+xml', f'Workbook MIME mismatch: {effective}'
        assert overrides.get('/xl/vbaProject.bin', defaults.get('bin')) == 'application/vnd.ms-office.vbaProject'
        assert archive.read('xl/vbaProject.bin')[:8] == bytes.fromhex('d0cf11e0a1b11ae1')
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        macro_rels = [r for r in rels if r.get('Type') == 'http://schemas.microsoft.com/office/2006/relationships/vbaProject']
        assert len(macro_rels) == 1 and macro_rels[0].get('Target') in ('vbaProject.bin', '/xl/vbaProject.bin')
    print('PASS: effective XLSM workbook MIME, VBA content type/relationship, ZIP and OLE signatures.')

workbook_path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'templates/ASC_Fail_Address_Fast.xlsm'
verify_macro_package(workbook_path)
parser = VBA_Parser(str(workbook_path))
sources = {name: code for _, _, name, code in parser.extract_macros()}
assert len(sources) == 4, list(sources)
code = next(code for name, code in sources.items() if name.startswith('ImportFailAddresses.'))
expected = (ROOT / 'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
assert code.replace('\r\n', '\n').strip() == expected.replace('\r\n', '\n').strip()
assert 'Public Sub ImportASCFailAddresses()' in code
assert 'Private Function TryParseAddress' in code
assert 'Private Sub FlushImportLog' not in code
assert 'Private Sub AddImportLog' not in code
assert 'Private mTokenRegex As Object' in code
assert 'ByRef bits() As Byte' in code
for name, source in sources.items():
    if not name.startswith('ImportFailAddresses.'):
        assert ' Sub ' not in source, name
parser.close()
print('PASS: independent olevba extraction; exact optimized BAS, byte-array flags, no log routines or sample event handlers.')
