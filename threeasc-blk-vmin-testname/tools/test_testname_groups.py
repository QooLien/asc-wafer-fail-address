"""Verify actual group selection, Vmin isolation and Excel-safe sheet naming."""
from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
source=(root/'src/BLKVminEngine.bas').read_text()
helpers=source[source.index('Private Function BuildNameGroups'):source.index('Private Function GroupDUTs')]
helpers+=source[source.index('Private Function IsNameSheet'):source.index('Private Sub SetupNameSheets')]
helpers+=source[source.index('Private Function BuildDUTGroup'):source.index('Private Function SortedDUTs')]
assert 'ws.Range("A2:C901").Value2, True' in source
assert 'ws.Range("D2:F901").Value2, False' in source
assert 'names(count) = Trim$(CStr(data(i, 2)))' in source
assert 'volts(count) = CDbl(data(i, 3))' in source
def adapt(text):
    text=text.replace('Private Function','Function').replace('Private Sub','Sub')
    text=re.sub(r'ByRef (names|allTNOs|allVolts|tnos|volts|cache|present)\(\)',r'ByRef \1',text)
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object)\b','',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text)
    text=text.replace('(1 To ','(').replace(', 1 To ',', ')
    for name in ('names','allTNOs','tnos','present'):text=text.replace(f'LBound({name})','1')
    return re.sub(r'\bNext \w+','Next',text)
tests=r'''
Dim names(4),allTNOs(4),allVolts(4),groups,tnos(),volts(),data,values(),result,i,b,used,a,c,saved
names(1)="W-Low pBL":names(2)="R-Low pBL":names(3)="W-Low pBL":names(4)=" r-low pBL "
allTNOs(1)="103":allTNOs(2)="101":allTNOs(3)="104":allTNOs(4)="102"
allVolts(1)=0.60:allVolts(2)=0.65:allVolts(3)=0.70:allVolts(4)=0.75
Set groups=BuildNameGroups(names)
If groups.Count<>2 Then Die "Trim/case grouping"
a=groups.Keys
If a(0)<>"W-Low pBL" Or a(1)<>"R-Low pBL" Then Die "First appearance group order"
Set data=CreateObject("Scripting.Dictionary")
For i=1 To 4
 ReDim values(255)
 For b=0 To 255:values(b)=CDbl(0):Next
 Select Case i
  Case 1:values(1)=2
  Case 2:values(0)=100
  Case 3:values(0)=2:values(1)=3
  Case 4:values(0)=3:values(1)=5
 End Select
 data.Add allTNOs(i) & "|8|12",values
Next
data.Add "101|9|12",values
GroupSequence allTNOs,allVolts,names,"W-Low pBL",tnos,volts
If UBound(tnos)<>2 Or tnos(1)<>"103" Or tnos(2)<>"104" Or volts(1)<>0.60 Or volts(2)<>0.70 Then Die "Interleaved W order / voltage"
If GroupHasDUT(data,tnos,"9|12") Then Die "Other-name DUT leaked"
result=BuildDUTGroup(data,tnos,volts,"8|12",0)
If result(3,5)<>0.60 Then Die "Legacy used voltage from another Test Name"
result=BuildDUTGroup(data,tnos,volts,"8|12",5)
If result(3,5)<>"Pass" Or result(3,6)<>0.60 Then Die "Threshold accumulated another Test Name"
If result(1,1)<>103 Or result(2,1)<>104 Or result(1,5)<>0 Or result(2,5)<>2 Or result(3,260)<>"Pass" Then Die "Grouped BLK mapping"
GroupSequence allTNOs,allVolts,names,"R-Low pBL",tnos,volts
If tnos(1)<>"101" Or tnos(2)<>"102" Or volts(1)<>0.65 Then Die "R order / voltage"
If Not GroupHasDUT(data,tnos,"9|12") Then Die "Own-name DUT lost"
result=BuildDUTGroup(data,tnos,volts,"8|12",5)
If result(3,5)<>"Fail" Or result(3,6)<>0.65 Then Die "R group threshold"
Set used=CreateObject("Scripting.Dictionary"):used.CompareMode=vbTextCompare
a=DistributionSheetName("W-Low pBL",used)
If a<>"BLK_Distribution(W-Low pBL)" Then Die "Requested sheet name"
a=DistributionSheetName("ABCDEFGHIJKLM-one",used):c=DistributionSheetName("ABCDEFGHIJKLM-two",used)
If Len(a)>31 Or Len(c)>31 Or a=c Then Die "Long-name collision"
a=DistributionSheetName("Bad[]:/\*?name",used)
For Each c In Array("[","]",":","/","\","*","?")
 If InStr(a,c)>0 Then Die "Invalid sheet character"
Next
If Not IsNameSheet(a) Then Die "Managed sheet recognition"
names(1)=" "
On Error Resume Next
Set groups=BuildNameGroups(names):saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Blank Test Name accepted"
WScript.Echo "PASS Test Name groups: ABC configuration mapping, interleaved/case-trim groups preserve TNO and voltage order; independent Legacy/cumulative Vmin; no cross-name DUT or FBC leakage; E:IZ mapping; 31-character collision-safe sheet names."
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
    p=Path(folder)/'names.vbs';p.write_text(adapt(tests+'\n'+helpers),encoding='ascii')
    result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
    output=(result.stdout+result.stderr).decode('cp950',errors='replace');print(output.encode('ascii','backslashreplace').decode())
    if 'PASS Test Name groups:' not in output:raise RuntimeError('Test Name tests did not complete')
