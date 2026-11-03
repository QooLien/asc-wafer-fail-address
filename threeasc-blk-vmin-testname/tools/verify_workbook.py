"""Independent final VBA/source, no-split, workbook and BLK mapping verification."""
from pathlib import Path
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from oletools.olevba import VBA_Parser
root=Path(__file__).resolve().parents[1]
path=Path(sys.argv[1]) if len(sys.argv)>1 else root/'templates/ASC_ThreeASC_Monitor_Function.xlsm'
main=(root/'src/ImportThreeASC_BLKVmin.bas').read_text()
engine=(root/'src/BLKVminEngine.bas').read_text()
baseline=(root/'reference/ImportThreeASC.bas').read_text()
assert main[main.index('Private Function PhaseTNOs'):]==baseline[baseline.index('Private Function PhaseTNOs'):]
def output_loop(s):
    a=s.index('        For Each v In addresses');return s[a:s.index('        Next v',a)+len('        Next v')]
assert output_loop(main)==output_loop(baseline)
assert 'SplitWTFailAddresses' not in main+engine and 'SplitWTByTNOSum' not in main+engine
assert 'ClassifyWTRow' not in main+engine and 'Private Function FailWT' not in main+engine
assert max(map(len,(main+engine).splitlines()))<1023
v=VBA_Parser(str(path));modules={n:c for _,_,n,c in v.extract_macros()}
assert len(modules)==7,list(modules)
for prefix,expected in [('ImportFailAddresses.',main),('BLKVminEngine.',engine)]:
    code=next(c for n,c in modules.items() if n.startswith(prefix))
    assert code.replace('\r\n','\n').strip()==expected.strip()
for name,code in modules.items():
    if not name.startswith(('ImportFailAddresses.','BLKVminEngine.')):assert ' Sub ' not in code
v.close()
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(path) as z:
    assert z.testzip() is None
    names=[n.get('name') for n in ET.fromstring(z.read('xl/workbook.xml')).find(S+'sheets')]
    assert names==['Config','Fail Address','BLK_Distribution(W-Low pBL)','BLK_Distribution(R-Low pBL)']
    strings=[]
    if 'xl/sharedStrings.xml' in z.namelist():strings=[''.join(n.itertext()) for n in ET.fromstring(z.read('xl/sharedStrings.xml'))]
    def value(c):
        if c.get('t')=='s':return strings[int(c.find(S+'v').text)]
        if c.get('t')=='inlineStr':return ''.join(c.find(S+'is').itertext())
        n=c.find(S+'v');return n.text if n is not None else ''
    def col(n):
        s=''
        while n:n,r=divmod(n-1,26);s=chr(65+r)+s
        return s
    config=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    cfg={c.get('r'):value(c) for c in config.iter(S+'c')}
    assert [cfg['A1'],cfg['B1'],cfg['C1']]==['TNO (Monitor)','Func.','Test Voltage']
    assert [cfg['D1'],cfg['E1'],cfg['F1']]==['TNO (Function)','Func.','Test Voltage']
    for r in range(2,902):
        assert all(cfg.get(f'{c}{r}','')=='' for c in 'DEF')
    assert 'ReadItemColumns ws.Range("D2:F901")' in engine
    assert 'Set tnos = BLKVminEngine.BLK_AddressTNOs()' in main
    assert 'BuildItemDUTGroup(mData, tnos, volts' in engine
    assert 'ws.Range("Q2")' in engine and 'cfg.Range("T2")' in main
    volts=[1.20,1.035,.99,.90,.81,.765,.72,.68,.64,.60,.56]
    defaults=[(350+i,'W-Low pBL',v) for i,v in enumerate(volts)]+[(370+i,'R-Low pBL',v) for i,v in enumerate(volts)]
    for r,(tno,name,voltage) in enumerate(defaults,2):
        assert int(cfg[f'A{r}'])==tno and cfg[f'B{r}']==name and float(cfg[f'C{r}'])==voltage
    assert len(defaults)==22
    for index in range(3,5):
        page=ET.fromstring(z.read(f'xl/worksheets/sheet{index}.xml'))
        row=next(r for r in page.find(S+'sheetData') if r.get('r')=='1')
        labels={c.get('r'):value(c) for c in row}
        for block in range(256):assert labels[col(block+7)+'1']=='BLK '+str(block)
        cells={c.get('r'):value(c) for c in page.iter(S+'c')}
        for wt in range(1,4):assert cells[f'F{wt+1}']==f'WT{wt}'
    page=ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
    labels={c.get('r'):value(c) for c in next(r for r in page.find(S+'sheetData') if r.get('r')=='1')}
    assert page.find(S+'autoFilter').get('ref')=='A1:BT1'
    for wt in range(1,4):
        for i,(tno,_,_) in enumerate(defaults):assert labels[col(7+(wt-1)*22+i)+'1']==f'WT{wt}\n{tno}'
    vml=ET.fromstring(z.read('xl/drawings/vmlDrawing1.vml'));X='{urn:schemas-microsoft-com:office:excel}'
    assert vml.find('.//'+X+'FmlaMacro').text=='[0]!ImportASCFailAddresses'
    anchor=[int(t.strip()) for t in vml.find('.//'+X+'Anchor').text.split(',')]
    assert anchor[0]==9 and anchor[2]==12
print('PASS: exact embedded VBA unchanged; all 22 default TNO/Name/Voltage pairs and 66 WT/TNO headers correct; two Test Name pages, 512 BLK labels, native button.')
