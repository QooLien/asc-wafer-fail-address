"""Verify the delivered classified workbook independently of the VBA writer."""
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from oletools.olevba import VBA_Parser
root=Path(__file__).resolve().parents[1]
source=(root/'src/ImportThreeASC_Classified.bas').read_text()
base=(root/'src/ImportThreeASC.bas').read_text()
assert source[source.index('Private Function PhaseTNOs'):]==base[base.index('Private Function PhaseTNOs'):]
assert source.index('If conflictCount > 0 Then Err.Raise') < source.index('dst.UsedRange.ClearContents')
path=root/'templates/ASC_Fail_Address_Classified.xlsm'
v=VBA_Parser(str(path));modules={n:c for _,_,n,c in v.extract_macros()}
assert len(modules)==6
code=next(c for n,c in modules.items() if n.startswith('ImportFailAddresses.'))
assert code.replace('\r\n','\n').strip()==source.strip()
for name,code in modules.items():
    if not name.startswith('ImportFailAddresses.'):assert ' Sub ' not in code
v.close()
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(path) as z:
    assert z.testzip() is None
    book=ET.fromstring(z.read('xl/workbook.xml'))
    assert [n.get('name') for n in book.find(S+'sheets')]==['Config','WT3 Fail','WT2 Fail','WT1 Fail']
    styles=ET.fromstring(z.read('xl/styles.xml'));xfs=styles.find(S+'cellXfs')
    for index in range(2,5):
        page=ET.fromstring(z.read(f'xl/worksheets/sheet{index}.xml'))
        assert page.find(S+'autoFilter').get('ref')=='A1:I1'
        for cell in page.iter(S+'c'):
            if cell.get('r') in ('G1','H1','I1'):
                assert xfs[int(cell.get('s'))].find(S+'alignment').get('textRotation')=='90'
    vml=ET.fromstring(z.read('xl/drawings/vmlDrawing1.vml'))
    X='{urn:schemas-microsoft-com:office:excel}'
    assert vml.find('.//'+X+'FmlaMacro').text=='[0]!ImportASCFailAddresses'
    assert '選擇 3 份 ASC 並分類' in z.read('xl/drawings/vmlDrawing1.vml').decode()
    assert vml.find('.//'+X+'Anchor').text.split(',')[2].strip()=='10'
print('PASS: original reader/parser/phase mapping unchanged; conflict check precedes output clear; exact embedded source, six modules, four sheets, rotated headers, filters and non-overlapping native button.')
