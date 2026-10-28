"""Test actual three-file header/index/parser/output code using independent fixture truth."""
from pathlib import Path
import re, subprocess, tempfile
root=Path(__file__).resolve().parents[1]
source=(root/'src/ImportThreeASC.bas').read_text()
base=(root/'src/ImportFailAddresses.bas').read_text()
assert source[source.index('Private Sub DetectEncoding'):]==base[base.index('Private Sub DetectEncoding'):]
def adapt(text):
    text=text.replace('Private Function','Function').replace('Private Sub','Sub')
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object|Byte)\b','',text)
    text=re.sub(r'ByRef (bits|nums)\(\)',r'ByRef \1',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text)
    text=re.sub(r'(\d)#',r'\1',text)
    text=text.replace('(0 To ','(').replace(', 0 To ',', ').replace('(1 To ','(').replace(', 1 To ',', ')
    return re.sub(r'\bNext \w+','Next',text)
helpers=source[source.index('Private Function PhaseTNOs'):source.index('Private Function ReadFirstMetadata')]
helpers+=source[source.index('Private Sub ParseASCBuffer'):]
a=source.index('    For phase = 1 To 3\n        Set localTNOs = PhaseTNOs(tnos, phase)',source.index('ReDim headers'))
header=source[a:source.index('    Next phase',a)+len('    Next phase')]
a=source.index('        For Each v In addresses')
output=source[a:source.index('        Next v',a)+len('        Next v')]
output=re.sub(r'(?m)^\s*dst\.Cells\(written \+ 2.*?Value2 = data$','                CaptureBatch data',output)
sort_start=source.index('    For phase = 1 To 2')
sorting=source[sort_start:source.index('    outputBase =',sort_start)]
check_start=source.index('        If phase > 1 Then')
identity=source[check_start:source.index('        End If',check_start)+len('        End If')]
tests=r'''
Const MAX_PENDING_RECORD=16384
Const MAX_RESULT_ROWS=1048575
Dim mTokenRegex,mBitMasks,tnos,localTNOs,addresses,bits(),flagBytes,capacity,phase,i,t
Dim buffer,headerTNO,sections,accepted,skipped,expected,key,addr,sizes,sz,p,raw
Dim headers(),columnCount,totalTNOs,data(),batchLimit,batchSize,batchRow,written,uniqueCount
Dim fields,v,rowIndex,b,bitValue,k,j,captured,variantCase,saved,dummy,chunkSizes
Dim paths(3),meta(3),steps(3),other,swapPath,swapStep,swapMeta
Set mTokenRegex=Nothing
mBitMasks=Array(1,2,4,8,16,32,64,128)
For Each variantCase In Array(1,3,9,900)
 Set tnos=CreateObject("Scripting.Dictionary")
 If variantCase=900 Then
  For i=0 To 899: tnos.Add CStr(100+((i*137) Mod 900)),tnos.Count: Next
 Else
  sizes=Array(999,100,777,108,255,512,101,888,107)
  For i=0 To variantCase-1: tnos.Add CStr(sizes(i)),tnos.Count: Next
 End If
 If variantCase=900 Then
  chunkSizes=Array(262144)
 Else
  chunkSizes=Array(7,262144)
 End If
 For Each sz In chunkSizes
  Set addresses=CreateObject("Scripting.Dictionary")
  Set expected=CreateObject("Scripting.Dictionary")
  totalTNOs=3*tnos.Count: flagBytes=(totalTNOs+7)\8: capacity=2
  ReDim bits(flagBytes-1,capacity-1)
  For phase=1 To 3
   Set localTNOs=PhaseTNOs(tnos,phase)
   buffer="": headerTNO="": sections=0: accepted=0: skipped=0: raw=""
   For i=999 To 100 Step -1
    If tnos.Exists(CStr(i)) Or i=500 Then
    addr="10,14,0," & ((i+phase) Mod 4) & ",2,3"
    raw=raw & "tno=" & i & " > [2," & i & "," & addr & ",]" & vbLf
    If tnos.Exists(CStr(i)) Then expected(addr & "|WT" & phase & vbLf & i)=True
    End If
   Next
   addr="10,14,0," & (20+phase) & ",2,3"
   raw=raw & "tno=999 > [2,999," & addr & ",]" & vbLf
   expected(addr & "|WT" & phase & vbLf & "999")=True
   For p=1 To Len(raw) Step sz
    buffer=buffer & Mid(raw,p,sz)
    ParseASCBuffer buffer,False,headerTNO,sections,accepted,skipped,localTNOs,addresses,bits,flagBytes,capacity
   Next
   ParseASCBuffer buffer,True,headerTNO,sections,accepted,skipped,localTNOs,addresses,bits,flagBytes,capacity
  Next
  columnCount=6+totalTNOs
  ReDim headers(1,columnCount)
__HEADER__
  uniqueCount=addresses.Count: batchLimit=3: batchSize=3: batchRow=0: written=0
  If batchSize>uniqueCount Then batchSize=uniqueCount
  ReDim data(batchSize,columnCount)
  Set captured=CreateObject("Scripting.Dictionary")
__OUTPUT__
  If captured.Count<>uniqueCount Or written<>uniqueCount Then Die "Batch row count"
  For Each key In expected
   If Not captured.Exists(Split(key,"|")(0)) Then Die "Missing union address"
  Next
 Next
 WScript.Echo "PASS three-group fixture: selected TNO count=" & variantCase
Next
If StepNumber("WTT0100")<>100 Or StepNumber("WTT2A")<>2 Or NormalWafer("#07")<>"07" Or NormalWafer("7")<>"07" Then Die "STEP/WNO normalization"
For phase=1 To 3
 Set meta(phase)=CreateObject("Scripting.Dictionary")
 meta(phase).Add "LOT","FAX7680T0"
 meta(phase).Add "WNO","07"
 meta(phase).Add "STEP",Array("WTT30","WTT2","WTT100")(phase-1)
 paths(phase)=CStr(meta(phase)("STEP"))
 steps(phase)=StepNumber(paths(phase))
Next
SortSources
If paths(1)<>"WTT2" Or paths(2)<>"WTT30" Or paths(3)<>"WTT100" Then Die "STEP numeric sorting"
For phase=1 To 3
 If paths(phase)<>meta(phase)("STEP") Then Die "Sorted metadata detached from file"
Next
phase=2: meta(2)("WNO")="7"
CheckIdentity
meta(2)("LOT")="FAX7681T0"
On Error Resume Next
CheckIdentity: saved=Err.Number: Err.Clear
On Error GoTo 0
If saved=0 Then Die "Full LOT mismatch with same first six characters not rejected"
meta(2)("LOT")="FAX7680T0": meta(2)("WNO")="8"
On Error Resume Next
CheckIdentity: saved=Err.Number: Err.Clear
On Error GoTo 0
If saved=0 Then Die "Wafer mismatch not rejected"
steps(1)=10: steps(2)=10: steps(3)=20
On Error Resume Next
SortSources: saved=Err.Number: Err.Clear
On Error GoTo 0
If saved=0 Then Die "Duplicate STEP not rejected"
For Each dummy In Array("WTTAAAA","WTT10A20","WTT2147483648")
 On Error Resume Next
 t=StepNumber(dummy): saved=Err.Number: Err.Clear
 On Error GoTo 0
 If saved=0 Then Die "Invalid STEP accepted"
Next
WScript.Echo "PASS: original reader/parser helpers unchanged; 3-group header and actual output loops; independent address/WT/TNO truth; 1/3/9/900 shuffled selections; 7-byte and large chunks; phase-only addresses, shared addresses, byte/group boundaries, full/partial batches; STEP/WNO validation."
Sub CaptureBatch(packet)
 Dim rr,cc,localKey,want,got
 For rr=1 To UBound(packet,1)
  localKey=""
  For cc=1 To 6
   If cc>1 Then localKey=localKey & ","
   localKey=localKey & packet(rr,cc)
  Next
  If captured.Exists(localKey) Then Die "Duplicate output address"
  captured.Add localKey,True
  For cc=7 To columnCount
   want=expected.Exists(localKey & "|" & headers(1,cc))
   got=Not IsEmpty(packet(rr,cc))
   If want<>got Then Die "Wrong WT/TNO column: " & headers(1,cc) & " at " & localKey
   If got Then
    If packet(rr,cc)<>1 Then Die "Non-1 output"
   End If
  Next
 Next
End Sub
Sub Die(message)
 WScript.Echo "FAIL: " & message: WScript.Quit 1
End Sub
Sub SortSources
__SORT__
End Sub
Sub CheckIdentity
__IDENTITY__
End Sub
'''
tests=tests.replace('__HEADER__',header).replace('__OUTPUT__',output).replace('__SORT__',sorting).replace('__IDENTITY__',identity)
with tempfile.TemporaryDirectory() as folder:
    p=Path(folder)/'three_test.vbs'
    p.write_text(adapt(tests+'\n'+helpers),encoding='ascii')
    subprocess.run(['cscript.exe','//nologo',str(p)],check=True)
