"""Test actual shared streaming feed plus original address parser; no Excel COM."""
from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
engine=(root/'src/BLKVminEngine.bas').read_text()
main=(root/'src/ImportThreeASC_BLKVmin.bas').read_text()
old=(root/'reference/ImportThreeASC.bas').read_text()
assert main[main.index('Private Function PhaseTNOs'):]==old[old.index('Private Function PhaseTNOs'):]
assert main.index('BLKVminEngine.BLK_Feed raw, False')<main.index('raw = ""',main.index('BLKVminEngine.BLK_Feed raw, False'))
helpers=engine[engine.index('Public Sub BLK_BeginFile'):engine.index('Public Function BLK_WritePhase')]
# VBScript has no On Error GoTo label; preserve feed body and emulate its catch after parser calls.
helpers=helpers[:helpers.index('Failed:')]+'End Sub\n'
helpers=helpers.replace('On Error GoTo Failed','On Error Resume Next')
lines=[]
for line in helpers.splitlines():
    lines.append(line)
    if 'ParseLine ' in line or 'FinishRecord:' in line:
        lines.append('    If Err.Number <> 0 Then mBLKError = "FBC Parse Error: " & Err.Description: Err.Clear: mCarry = "": mActive = False: Exit Sub')
helpers='\n'.join(lines)+'\n'
helpers+=engine[engine.index('Private Sub ResetParser'):engine.index('Private Function SortedDUTs')]
helpers+=engine[engine.index('Private Function UnsignedValue'):engine.index('Private Sub WriteGroup')]
helpers+=main[main.index('Private Sub ParseASCBuffer'):]
def adapt(text):
    for kind in ('Private','Public'):
        text=text.replace(kind+' Sub','Sub').replace(kind+' Function','Function')
    text=re.sub(r'ByRef (cache|present|volts|tnos|bits|nums)\(\)',r'ByRef \1',text)
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object|Byte)\b','',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text);text=re.sub(r'(\d)#',r'\1',text)
    text=text.replace('(0 To ','(').replace(', 0 To ',', ').replace('(1 To ','(').replace(', 1 To ',', ')
    text=text.replace('LBound(present)','1').replace('LBound(mSequence)','1')
    return re.sub(r'\bNext \w+','Next',text)
tests=r'''
Const MAX_EXACT=9007199254740991
Const MAX_LINE=1048576
Const MAX_PENDING_RECORD=16384
Const MAX_RESULT_ROWS=1048575
Dim mTags,mTagHeads,mBlockTag,mSpaces,mDigits,mData,mWanted,mDUTs,mTno,mX,mY,mPath,mActive
Dim mRecordKey,mRecordDUT,mValues(),mCount,mLine,mRecordLine,mSequence(4),mCarry,mBLKError
Dim mTokenRegex,mBitMasks,tnos,addresses,bits(),flagBytes,capacity,headerTNO,sections,accepted,skipped,buffer
Dim phase,i,b,p,chunk,raw,values(),item,key,index,wanted,actual,seq
Dim fso,volts(4),group,stepValue,expected255
Set mTags=Nothing:Set mDigits=Nothing:Set mTokenRegex=Nothing
mBitMasks=Array(1,2,4,8,16,32,64,128)
seq=Array(103,101,104,102)
For i=1 To 4:mSequence(i)=CStr(seq(i-1)):Next
For Each chunk In Array(7,262144)
 Set addresses=CreateObject("Scripting.Dictionary")
 flagBytes=2:capacity=2:ReDim bits(1,1)
 For phase=1 To 3
  Set tnos=CreateObject("Scripting.Dictionary")
  For i=1 To 4:tnos.Add mSequence(i),(phase-1)*4+i-1:Next
  BLK_BeginFile "WT" & phase
  raw="LOT=FAX7680T0 WNO=7 STEP=WTT" & phase & " Program=VER_FA" & vbLf
  For i=1 To 4
   ReDim values(255)
   For b=0 To 255:values(b)=phase*1000+i*100+b:Next
   raw=raw & "Tno=" & mSequence(i) & " X=8 Y=12 FBC BY BLK=" & Join(values," ") & vbLf
   raw=raw & "[2," & mSequence(i) & ",8,12,0,0,0,0,] [2," & mSequence(i) & ",8,12,0," & phase & ",0,0,]" & vbLf
  Next
  buffer="":headerTNO="":sections=0:accepted=0:skipped=0
  For p=1 To Len(raw) Step chunk
   item=Mid(raw,p,chunk)
   BLK_Feed item,False
   buffer=buffer & item
   ParseASCBuffer buffer,False,headerTNO,sections,accepted,skipped,tnos,addresses,bits,flagBytes,capacity
  Next
  BLK_Feed "",True
  ParseASCBuffer buffer,True,headerTNO,sections,accepted,skipped,tnos,addresses,bits,flagBytes,capacity
  If Len(mBLKError)>0 Or mData.Count<>4 Or accepted<>8 Then Die "Shared stream parsing mismatch: " & mBLKError
  For i=1 To 4
   item=mData(mSequence(i) & "|8|12")
   For b=0 To 255
    If item(b)<>phase*1000+i*100+b Then Die "Cross-WT FBC contamination"
   Next
  Next
 Next
 If addresses.Count<>4 Then Die "Original address union changed"
 For Each key In addresses
  item=Split(key,",")
  For phase=1 To 3
   For i=1 To 4
    index=(phase-1)*4+i-1
    actual=((bits(index\8,addresses(key)) And mBitMasks(index And 7))<>0)
    wanted=(CLng(item(3))=0 Or CLng(item(3))=phase)
    If actual<>wanted Then Die "Original WT/TNO flag shifted"
   Next
  Next
 Next
Next
BLK_BeginFile "address only"
BLK_Feed "tno=103 > [2,103,8,12,0,0,0,0,]" & vbLf,True
If Len(mBLKError)>0 Or mData.Count<>0 Then Die "Address-only file not handled as No Data"
BLK_BeginFile "bad FBC"
ReDim values(254):For b=0 To 254:values(b)="0":Next
raw="Tno=103 X=8 Y=12 FBC BY BLK=" & Join(values," ") & vbLf & "[2,103,8,12,0,0,0,0,]" & vbLf
BLK_Feed raw,True
If Len(mBLKError)=0 Then Die "Malformed FBC not isolated"
Set tnos=CreateObject("Scripting.Dictionary"):tnos.Add "103",0
Set addresses=CreateObject("Scripting.Dictionary"):capacity=2:flagBytes=1:ReDim bits(0,1)
buffer=raw:headerTNO="":sections=0:accepted=0:skipped=0
ParseASCBuffer buffer,True,headerTNO,sections,accepted,skipped,tnos,addresses,bits,flagBytes,capacity
If accepted<>1 Then Die "FBC error stopped original address parser"
BLK_BeginFile "next WT"
If Len(mBLKError)>0 Then Die "FBC error leaked across files"
Set fso=CreateObject("Scripting.FileSystemObject")
volts(1)=0.60:volts(2)=0.65:volts(3)=0.70:volts(4)=0.75
Set addresses=CreateObject("Scripting.Dictionary"):capacity=2:flagBytes=2:ReDim bits(1,1)
For phase=1 To 3
 stepValue=Array(2,30,100)(phase-1)
 raw=fso.OpenTextFile(fso.BuildPath(WScript.Arguments(0),"WT" & phase & "_STEP" & stepValue & ".asc")).ReadAll
 BLK_BeginFile "example WT" & phase
 BLK_Feed raw,True
 If Len(mBLKError)>0 Then Die "Example FBC parse: " & mBLKError
 group=BuildDUTGroup(mData,mSequence,volts,"8|12",0)
 If group(5,5)<>0.65 Or group(5,260)<>"Fail" Then Die "Example Legacy judge"
 group=BuildDUTGroup(mData,mSequence,volts,"8|12",5)
 expected255=Array("Pass",0.65,0.60)(phase-1)
 If group(5,5)<>0.70 Or group(5,6)<>"Pass" Or CStr(group(5,260))<>CStr(expected255) Then Die "Example cumulative judge"
 Set tnos=CreateObject("Scripting.Dictionary")
 For i=1 To 4:tnos.Add mSequence(i),(phase-1)*4+i-1:Next
 buffer=raw:headerTNO="":sections=0:accepted=0:skipped=0
 ParseASCBuffer buffer,True,headerTNO,sections,accepted,skipped,tnos,addresses,bits,flagBytes,capacity
 If accepted<>phase*4+8 Then Die "Example address/FBC totals differ"
Next
If addresses.Count<>9 Then Die "Example original address union"
WScript.Echo "PASS shared decoded stream: mixed address/FBC data; 7/262144 chunks; three WT isolation; all 256 FBC values; original address union/WT/TNO flags unchanged; missing FBC and FBC errors do not stop address parsing."
WScript.Echo "PASS delivered three-ASC fixtures: 9 unique addresses, per-file FBC/address counts equal, independent Legacy/Threshold5 Vmin including BLK255."
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
    p=Path(folder)/'shared.vbs';p.write_text(adapt(tests+'\n'+helpers),encoding='ascii')
    result=subprocess.run(['cscript.exe','//nologo',str(p),str(root/'examples/three_asc')],capture_output=True)
    output=(result.stdout+result.stderr).decode('cp950',errors='replace');print(output.encode('ascii','backslashreplace').decode())
    if 'PASS delivered three-ASC fixtures:' not in output:
        match=re.search(r'\((\d+),\s*\d+\)',output)
        if match:
            n=int(match.group(1));lines=p.read_text().splitlines()
            for i in range(max(0,n-2),min(len(lines),n+1)):print(f'{i+1}: {lines[i]}')
        raise RuntimeError('Shared feed test did not complete')
