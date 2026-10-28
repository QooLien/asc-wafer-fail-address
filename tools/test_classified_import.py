"""Execute actual classification and grouped output with independent address/WT/TNO truth."""
from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
source=(root/'src/ImportThreeASC_Classified.bas').read_text()
base=(root/'src/ImportThreeASC.bas').read_text()
assert source[source.index('Private Function PhaseTNOs'):]==base[base.index('Private Function PhaseTNOs'):]
def adapt(text):
    text=text.replace('Private Function','Function').replace('Private Sub','Sub')
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object|Byte)\b','',text)
    text=re.sub(r'ByRef (bits|nums)\(\)',r'ByRef \1',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text)
    text=re.sub(r'(\d)#',r'\1',text)
    text=text.replace('(0 To ','(').replace(', 0 To ',', ').replace('(1 To ','(').replace(', 1 To ',', ')
    return re.sub(r'\bNext \w+','Next',text)
helpers=source[source.index('Private Function FailWT'):source.index('Private Function StepNumber')]
helpers+=source[source.index('Private Sub ParseASCBuffer'):]
a=source.index('    stage = "Classify selected-TNO')
classification=source[a:source.index('    columnCount =',a)]
a=source.index('    For phase = 1 To 3\n        Set localTNOs',source.index('ReDim headers'))
header=source[a:source.index('    Next phase',a)+len('    Next phase')]
a=source.index('    If groupCounts(group) > 0 Then',source.index('stage = "Write output batches"'))
output=source[a:source.index('    stage = "Format output"',a)]
output=re.sub(r'(?m)^\s*dst\.Cells\(written \+ 2.*?Value2 = data$','                CaptureBatch data',output)
tests=r'''
Const MAX_PENDING_RECORD=16384
Const MAX_RESULT_ROWS=1048575
Dim mTokenRegex,mBitMasks,tnos,localTNOs,addresses,bits(),flagBytes,capacity,phase,i,t,group
Dim buffer,headerTNO,sections,accepted,skipped,expected,key,addr,sizes,raw,expectedGroups
Dim headers(),columnCount,totalTNOs,data(),batchLimit,batchSize,batchRow,written,uniqueCount
Dim fields,v,rowIndex,b,bitValue,k,j,captured,variantCase,groups(),groupCounts(3),judged
Dim conflictCount,conflictExamples,stage,checks,counts(3),p,index,mode
Set mTokenRegex=Nothing:mBitMasks=Array(1,2,4,8,16,32,64,128)
For Each variantCase In Array(1,3,9,900)
 Set tnos=CreateObject("Scripting.Dictionary")
 If variantCase=900 Then
  For i=0 To 899:tnos.Add CStr(100+((i*137) Mod 900)),tnos.Count:Next
 Else
  sizes=Array(100,777,108,255,512,101,888,107,998)
  For i=0 To variantCase-1:tnos.Add CStr(sizes(i)),tnos.Count:Next
 End If
 For Each mode In Array(0,3,4)
 Set addresses=CreateObject("Scripting.Dictionary")
 Set expected=CreateObject("Scripting.Dictionary")
 Set expectedGroups=CreateObject("Scripting.Dictionary")
 totalTNOs=3*tnos.Count:flagBytes=(totalTNOs+7)\8:capacity=2
 ReDim bits(flagBytes-1,capacity-1)
 For phase=1 To 3
  Set localTNOs=PhaseTNOs(tnos,phase)
  buffer="":headerTNO="":sections=0:accepted=0:skipped=0:raw="tno=100 >" & vbLf
  For i=999 To 100 Step -1
   If tnos.Exists(CStr(i)) And (mode=0 Or mode=phase) Then
    addr="10,14,0," & (phase*100+(i Mod 5)) & ",2,3"
    raw=raw & "tno=" & i & " > [2," & i & "," & addr & ",]" & vbLf
    expected(addr & "|WT" & phase & vbLf & i)=True
    expectedGroups(addr)=phase
   End If
  Next
  ' Unselected monitor failures in earlier WT must not move this WT3 address.
  If Not tnos.Exists("999") And phase<3 Then raw=raw & "tno=999 > [2,999,10,14,0,300,2,3,]" & vbLf
  For p=1 To Len(raw) Step 127
   buffer=buffer & Mid(raw,p,127)
   ParseASCBuffer buffer,False,headerTNO,sections,accepted,skipped,localTNOs,addresses,bits,flagBytes,capacity
  Next
  ParseASCBuffer buffer,True,headerTNO,sections,accepted,skipped,localTNOs,addresses,bits,flagBytes,capacity
 Next
 uniqueCount=addresses.Count:conflictCount=0:conflictExamples=""
 For i=1 To 3:groupCounts(i)=0:counts(i)=0:Next
__CLASSIFY__
 If uniqueCount<>expectedGroups.Count Then Die "Monitor-only address leaked or missing address"
 columnCount=6+totalTNOs:ReDim headers(1,columnCount)
__HEADER__
 Set captured=CreateObject("Scripting.Dictionary")
 For group=1 To 3
  written=0:batchRow=0:batchLimit=2
__OUTPUT__
  If written<>groupCounts(group) Then Die "Per-sheet batch row count"
 Next
 If captured.Count<>expectedGroups.Count Then Die "Classified union count"
 For Each key In expectedGroups
  If Not captured.Exists(key) Then Die "Address lost in classified output"
 Next
 Next
 WScript.Echo "PASS classified fixture: selected TNO count=" & variantCase & "; three groups / WT3 only / no fail"
Next
' Exhaust all seven nonempty WT presence combinations, with nonmatching TNOs in each phase.
For Each variantCase In Array(1,3,9,900)
 flagBytes=(3*variantCase+7)\8
 For checks=0 To 7
  ReDim bits(flagBytes-1,0)
  For phase=1 To 3
   If (checks And mBitMasks(phase-1))<>0 Then
    index=(phase-1)*variantCase+((phase-1) Mod variantCase)
    bits(index\8,0)=bits(index\8,0) Or mBitMasks(index And 7)
   End If
  Next
  judged=FailWT(bits,0,variantCase)
  Select Case checks
   Case 0:If judged<>0 Then Die "Empty address classified as fail"
   Case 1:If judged<>1 Then Die "WT1 judge"
   Case 2:If judged<>2 Then Die "WT2 judge"
   Case 4:If judged<>3 Then Die "WT3 judge"
   Case Else:If judged<>-1 Then Die "Cross-WT conflict not rejected"
  End Select
 Next
Next
WScript.Echo "PASS: actual classifier and three-sheet batch output; Config-only judgment, monitor exclusion; independent address/WT/TNO per-cell oracle; 1/3/9/900 selections; every WT-presence combination, including different TNO cross-WT conflict; full/partial batches."
Sub CaptureBatch(packet)
 Dim rr,cc,localKey,want,got
 For rr=1 To UBound(packet,1)
  localKey=""
  For cc=1 To 6
   If cc>1 Then localKey=localKey & ","
   localKey=localKey & packet(rr,cc)
  Next
  If captured.Exists(localKey) Then Die "Address duplicated across sheets"
  If expectedGroups(localKey)<>group Then Die "Wrong WT Fail sheet"
  captured.Add localKey,True
  For cc=7 To columnCount
   want=expected.Exists(localKey & "|" & headers(1,cc))
   got=Not IsEmpty(packet(rr,cc))
   If want<>got Then Die "Wrong TNO column or unexpected nonblank: " & localKey & " / " & headers(1,cc)
   If got Then
    If packet(rr,cc)<>1 Then Die "Expected numeric 1"
    If Left(headers(1,cc),3)<>"WT" & group Then Die "Other WT has 1 on this sheet"
   End If
  Next
 Next
End Sub
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
'''
tests=tests.replace('__CLASSIFY__',classification).replace('__HEADER__',header).replace('__OUTPUT__',output)
with tempfile.TemporaryDirectory() as folder:
    p=Path(folder)/'classified.vbs'
    p.write_text(adapt(tests+'\n'+helpers),encoding='ascii')
    subprocess.run(['cscript.exe','//nologo',str(p)],check=True)
