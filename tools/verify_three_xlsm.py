"""Independent embedded VBA and saved three-file workbook verification."""
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET
from oletools.olevba import VBA_Parser
root=Path(__file__).resolve().parents[1]
path=root/'templates/ASC_Fail_Address_ThreeASC.xlsm'
source=(root/'src/ImportThreeASC.bas').read_text()
base=(root/'src/ImportFailAddresses.bas').read_text()
assert source[source.index('Private Sub DetectEncoding'):]==base[base.index('Private Sub DetectEncoding'):]
def output_loop(text):
    start=text.index('        For Each v In addresses')
    return text[start:text.index('        Next v',start)+len('        Next v')]
assert output_loop(source)==output_loop(base)
assert max(map(len,source.splitlines()))<1023
v=VBA_Parser(str(path))
modules={name:code for _,_,name,code in v.extract_macros()}
assert len(modules)==4
embedded=next(code for name,code in modules.items() if name.startswith('ImportFailAddresses.'))
assert embedded.replace('\r\n','\n').strip()==source.strip()
for name,code in modules.items():
    if not name.startswith('ImportFailAddresses.'):
        assert ' Sub ' not in code
v.close()
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(path) as z:
    assert z.testzip() is None
    book=ET.fromstring(z.read('xl/workbook.xml'))
    assert [s.get('name') for s in book.find(S+'sheets')]==['Config','Fail Address']
    sheet=ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
    styles=ET.fromstring(z.read('xl/styles.xml'))
    xfs=styles.find(S+'cellXfs');fills=styles.find(S+'fills')
    colors=[]
    for cell in sheet.iter(S+'c'):
        if cell.get('r') not in ('G1','H1','I1'):continue
        xf=xfs[int(cell.get('s'))]
        assert xf.find(S+'alignment').get('textRotation')=='90'
        colors.append(fills[int(xf.get('fillId'))].find(S+'patternFill').find(S+'fgColor').get('rgb'))
    assert len(set(colors))==3
    assert sheet.find(S+'autoFilter').get('ref')=='A1:I1'
    vml=z.read('xl/drawings/vmlDrawing1.vml').decode()
    assert '[0]!ImportASCFailAddresses' in vml and '選擇 3 份 ASC 並合併' in vml
print('PASS: original helpers and data output loop unchanged; embedded three-ASC VBA exact source; four modules, two sheets, vertical WT/TNO headers, three colors, filter and native button.')
