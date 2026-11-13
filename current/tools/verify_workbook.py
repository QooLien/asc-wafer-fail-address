from pathlib import Path
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from oletools.olevba import VBA_Parser
r=Path(__file__).resolve().parents[1];p=Path(sys.argv[1]) if len(sys.argv)>1 else r/'templates/SRAM Extra BLK Failure  Analysis_v13.xlsm'
v=VBA_Parser(str(p));modules={n:c for _,_,n,c in v.extract_macros()};v.close()
assert len(modules)==11
for prefix,file in [('ImportFailAddresses.','ImportThreeASC_BLKVmin.bas'),('BLKVminEngine.','BLKVminEngine.bas')]:
 expected=(r/'src'/file).read_text().strip()
 actual=next(c for n,c in modules.items() if n.startswith(prefix)).replace('\r\n','\n').strip()
 assert actual==expected
book_source=next(c for n,c in modules.items() if n.startswith('ThisWorkbook.')).replace('\r\n','\n')
assert (r/'src/WorkbookMapEvents.vba').read_text().strip() in book_source
assert 'FixtureDoubleClick' not in ''.join(modules.values())
main=(r/'src/ImportThreeASC_BLKVmin.bas').read_text();base=(r/'reference/ImportThreeASC.bas').read_text()
assert 'NativeConsumeTestRaw' not in main
assert main[main.index('Private Function PhaseTNOs'):]==base[base.index('Private Function PhaseTNOs'):]
def loop(s):
 a=s.index('        For Each v In addresses');return s[a:s.index('        Next v',a)]
assert loop(main)==loop(base)
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(p) as z:
 assert z.testzip() is None
 book=ET.fromstring(z.read('xl/workbook.xml'))
 sheets={n.get('name'):n for n in book.find(S+'sheets')}
 assert set(sheets)=={'Config','Fail Address','FBC by BLK','BLK_FBC_Comparison','Chip Locations','BLK Fail Count','Missing BLK Address','Import Index','Log'}
 assert sheets['Import Index'].get('state')=='veryHidden'
 assert sheets['FBC by BLK'].get('state')=='veryHidden'
 rels={n.get('Id'):n.get('Target') for n in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
 R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
 parts={name:(rels[node.get(R+'id')].lstrip('/') if rels[node.get(R+'id')].startswith('/') else 'xl/'+rels[node.get(R+'id')]) for name,node in sheets.items()}
 for part in parts.values():
  sh=ET.fromstring(z.read(part))
  for view in sh.find(S+'sheetViews'):assert view.get('showGridLines')=='0'
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
 style_root=ET.fromstring(z.read('xl/styles.xml'))
 zero_cell=next(c for c in sheet.iter(S+'c') if c.get('r')=='E2')
 zero_xf=list(style_root.find(S+'cellXfs'))[int(zero_cell.get('s'))]
 fmts={n.get('numFmtId'):n.get('formatCode') for n in style_root.find(S+'numFmts')}
 assert fmts[zero_xf.get('numFmtId')].replace('\\-','-')=='0;-0;;@'
 for b in range(256):assert cells[col(b+5)+'1']=='BLK '+str(b)
 cfg=ET.fromstring(z.read('xl/worksheets/sheet1.xml'));cells={c.get('r'):value(c) for c in cfg.iter(S+'c')}
 defaults=[350+i for i in range(11)]+[370+i for i in range(11)]
 assert [int(cells[f'A{i+2}']) for i in range(22)]==defaults
 assert 'VminFBCThreshold' not in cells.values()
 assert cells['G1']=='Stop Sign'
 assert not cells.get('T2')
 log=ET.fromstring(z.read(parts['Log']))
 log_values={c.get('r'):value(c) for c in log.iter(S+'c')}
 assert [log_values[c+'1'] for c in 'ABCD']==['Time','Operation','Status','Message']
 # User Config may deliberately remove every information/control cell to the right of G.
 if cells.get('Q2'):
  assert cells['Q2']=='Fail Map'
  assert cfg.find(S+'dataValidations') is not None
 g1=next(c for c in cfg.iter(S+'c') if c.get('r')=='G1')
 styles=ET.fromstring(z.read('xl/styles.xml'))
 xf=list(styles.find(S+'cellXfs'))[int(g1.get('s'))]
 font=list(styles.find(S+'fonts'))[int(xf.get('fontId'))]
 fill=list(styles.find(S+'fills'))[int(xf.get('fillId'))]
 assert font.find(S+'color').get('rgb').endswith('FFFFFF')
 assert fill.find('.//'+S+'fgColor').get('rgb').endswith('C00000')
 fail=ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
 titles={c.get('r'):value(c) for c in fail.iter(S+'c')}
 metadata=ET.fromstring(z.read(parts['Import Index']))
 meta_values={c.get('r'):value(c) for c in metadata.iter(S+'c')}
 phases=sorted({v for ref,v in meta_values.items() if ref.startswith('Q') and ref!='Q1' and v.startswith('WT')})
 configured=[]
 for tc,fc,vc in [('A','B','C'),('D','E','F')]:
  for row in range(2,902):
   if cells.get(f'{tc}{row}'):
    label=cells[f'{tc}{row}'];func=cells.get(f'{fc}{row}','');voltage=cells.get(f'{vc}{row}','')
    if func and voltage:
     whole,frac=f'{float(voltage):.9f}'.split('.');label=func+'_'+whole+'.'+frac.rstrip('0').ljust(2,'0')+'v'
    configured.append(label)
 for phase,wt in enumerate(phases):
  for i,label in enumerate(configured):assert titles[col(7+phase*len(configured)+i)+'1']==wt+'\n'+label
 assert [s.get('name') for s in book.find(S+'sheets')][-2:]==['Chip Locations','Log']
 assert 90 <= float(fail.find(S+'sheetData').find(S+'row').get('ht')) <= 200
 vml=ET.fromstring(z.read('xl/drawings/vmlDrawing1.vml'))
 X='{urn:schemas-microsoft-com:office:excel}'
 buttons=vml.findall('.//'+X+'ClientData')
 macros={b.find(X+'FmlaMacro').text:b for b in buttons}
 assert set(macros)=={'[0]!ImportASCFailAddresses','[0]!ImportASCFBC','[0]!BuildFailMaps','[0]!FindMissingBLKAddresses'}
 for macro,row in [('[0]!ImportASCFailAddresses',12),('[0]!ImportASCFBC',15),('[0]!BuildFailMaps',18),('[0]!FindMissingBLKAddresses',21)]:
  anchor=[int(v.strip()) for v in macros[macro].find(X+'Anchor').text.split(',')]
  assert (anchor[0],anchor[2]) in [(9,row),(7,row-11)]
print('PASS XLSM: eleven exact VBA modules, nine sheets with visible BLK_FBC_Comparison, hidden raw/index, four original macro buttons, original address helpers/output and all256 FBC headers.')
