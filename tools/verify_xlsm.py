"""Independent VBA extraction check. Requires oletools (development only)."""
from pathlib import Path
from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
parser = VBA_Parser(str(ROOT / 'templates/ASC_Fail_Address.xlsm'))
sources = {name: code for _, _, name, code in parser.extract_macros()}
assert len(sources) == 6, list(sources)
code = next(code for name, code in sources.items() if name.startswith('ImportFailAddresses.'))
expected = (ROOT / 'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
assert code.replace('\r\n', '\n').strip() == expected.replace('\r\n', '\n').strip()
assert 'Public Sub ImportASCFailAddresses()' in code
assert 'Private Function TryParseAddress' in code
assert 'Private Sub FlushImportLog' in code
assert 'Private Sub AddImportLog' in code
for name, source in sources.items():
    if not name.startswith('ImportFailAddresses.'):
        assert ' Sub ' not in source, name
parser.close()
print('PASS: independent olevba extraction; all 6 modules, exact revised BAS source including log helpers, no sample event handlers.')
