from pathlib import Path
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from oletools.olevba import VBA_Parser
r=Path(__file__).resolve().parents[1];p=Path(sys.argv[1]) if len(sys.argv)>1 else r/'templates/ASC_ThreeASC_FBC.xlsm'
v=VBA_Parser(str(p));modules={n:c for _,_,n,c in v.extract_macros()};v.close()
assert len(modules)==6
for prefix,file in [('ImportFailAddresses.','ImportThreeASC_BLKVmin.bas'),('BLKVminEngine.','BLKVminEngine.bas')]:
 expected=(r/'src'/file).read_text().strip()
 actual=next(c for n,c in modules.items() if n.startswith(prefix)).replace('\r\n','\n').strip()
 assert actual==expected
main=(r/'src/ImportThreeASC_BLKVmin.bas').read_text();base=(r/'reference/ImportThreeASC.bas').read_text()
assert main[main.index('Private Function PhaseTNOs'):]==base[base.index('Private Function PhaseTNOs'):]
def loop(s):
 a=s.index('        For Each v In addresses');return s[a:s.index('        Next v',a)]
assert loop(main)==loop(base)
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(p) as z:
 assert z.testzip() is None
 book=ET.fromstring(z.read('xl/workbook.xml'))
 assert [n.get('name') for n in book.find(S+'sheets')]==['Config','Fail Address','FBC by BLK']
 strings=[]
 if 'xl/sharedStrings.xml' in z.namelist():strings=[''.join(n.itertext()) for n in ET.fromstring(z.read('xl/sharedStrings.xml'))]
 def value(c):
  if c.get('t')=='s':return strings[int(c.find(S+'v').text)]
  if c.get('t')=='inlineStr':return ''.join(c.find(S+'is').itertext())
  v=c.find(S+'v');return '' if v is None else v.text
 def col(n):
  text=''
  while n:n,d=divmod(n-1,26);text=chr(65+d)+text
  return text
 sheet=ET.fromstring(z.read('xl/worksheets/sheet3.xml'));cells={c.get('r'):value(c) for c in sheet.iter(S+'c')}
 assert [cells[c+'1'] for c in 'ABCD']==['WT','TNO','X','Y']
 for b in range(256):assert cells[col(b+5)+'1']=='BLK '+str(b)
 cfg=ET.fromstring(z.read('xl/worksheets/sheet1.xml'));cells={c.get('r'):value(c) for c in cfg.iter(S+'c')}
 defaults=[350+i for i in range(11)]+[370+i for i in range(11)]
 assert [int(cells[f'A{i+2}']) for i in range(22)]==defaults
 assert 'VminFBCThreshold' not in cells.values()
 labels=['1.20','1.035','0.99','0.90','0.81','0.765','0.72','0.68','0.64','0.60','0.56']
 fail=ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
 titles={c.get('r'):value(c) for c in fail.iter(S+'c')}
 for wt in range(1,4):
  for i in range(22):
   func='W-Low pBL' if i<11 else 'R-Low pBL'
   assert titles[col(7+(wt-1)*22+i)+'1']==f'WT{wt}\n{func}_{labels[i%11]}v'
 assert float(fail.find(S+'sheetData').find(S+'row').get('ht'))==150
 vml=ET.fromstring(z.read('xl/drawings/vmlDrawing1.vml'))
 X='{urn:schemas-microsoft-com:office:excel}'
 buttons=vml.findall('.//'+X+'ClientData')
 macros={b.find(X+'FmlaMacro').text:b for b in buttons}
 assert set(macros)=={'[0]!ImportASCFailAddresses','[0]!ImportASCFBC'}
 for macro,row in [('[0]!ImportASCFailAddresses',12),('[0]!ImportASCFBC',15)]:
  anchor=[int(v.strip()) for v in macros[macro].find(X+'Anchor').text.split(',')]
  assert anchor[0]==9 and anchor[2]==row
print('PASS XLSM: six exact VBA modules, only three sheets, unchanged address helpers/output, all256 E:IZ BLK headers and22 default TNOs, no Vmin configuration.')
