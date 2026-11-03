from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/BLKVminEngine.bas').read_text()
def part(start,end):return s[s.index(start):s.index(end,s.index(start))]
helpers=part('Private Sub ResetParser','Private Sub ParseLine')
helpers+=part('Private Function NormalizeASCWhitespace','Private Sub SortXY')
helpers+=part('Private Function Coordinate','Public Function BLK_RecordCount')
helpers+=part('Private Sub CaptureMapLocation','Private Sub WriteMapLocations')
for kind in ['Private','Public']:helpers=helpers.replace(kind+' Sub','Sub').replace(kind+' Function','Function')
helpers=re.sub(r' As (String|Variant|Long|Double|Object|Boolean)\b','',helpers)
helpers=re.sub(r'\b(\w+)\$',r'\1',helpers);helpers=re.sub(r'(\d)#',r'\1',helpers)
helpers=re.sub(r'\bNext \w+','Next',helpers)
code='''
Dim mData,mDUTs,mTno,mX,mY,mPath,mLine,mActive,mCount,mTags,mBlockTag,mSpaces,mDigits,mLocations
Set mTags=Nothing:Set mLocations=CreateObject("Scripting.Dictionary")
ResetParser "WT1"
CaptureMapLocation "DUT=9 X=10 Y=17 DC=Fail"
CaptureMapLocation "Tno=0 DUT=10 X=11 Y=17 ISB=Fail"
CaptureMapLocation "Tno=350 X=12 Y=17 FBC BY BLK=0 1 2"
CaptureMapLocation "Tno=501 X=12 Y=17 FBC BY BLK=0 1 2"
CaptureMapLocation "Instrument X=99 Y=99"
CaptureMapLocation "Tno= X= ChipX Y= ChipY"
CaptureMapLocation "Tno=501 X=13"
CaptureMapLocation "Tno=501 Y=18"
If mLocations.Count<>3 Or mLocations.Exists("13|18") Or mLocations.Exists("99|99") Then Die "Bad XY pairs / stale context"
ResetParser "WT2"
CaptureMapLocation "DUT=4 X=-1 Y=2"
If mLocations.Count<>4 Or Not mLocations.Exists("10|17") Or Not mLocations.Exists("-1|2") Then Die "Union across WTs lost"
WScript.Echo "PASS Raw Chip locations: DC/ISB/no-FBC and Monitor records supply coordinates only; union across WT files; unique actual XY pairs; no metadata or cross-line Cartesian locations."
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'locations.vbs';p.write_text(code+'\n'+helpers,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS Raw Chip locations:' in out,out
