"""Test actual no-log VBA parser and output batching in the VBScript engine.

Syntax adaptation only; Byte arrays become Variant arrays in VBScript, so this
does not measure native VBA memory use or Excel I/O performance.
"""
import argparse
import pathlib
import re
import subprocess
import tempfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
source = (ROOT / 'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
def adapt(text):
    text = text.replace('Private Function', 'Function').replace('Private Sub', 'Sub')
    text = re.sub(r' As (String|Variant|Long|Double|Boolean|Object|Byte)\b', '', text)
    text = re.sub(r'ByRef (bits|nums)\(\)', r'ByRef \1', text)
    text = re.sub(r'\b(\w+)\$', r'\1', text)
    text = re.sub(r'(\d)#', r'\1', text)
    text = text.replace('(0 To ', '(').replace(', 0 To ', ', ').replace('(1 To ', '(').replace(', 1 To ', ', ')
    return re.sub(r'\bNext \w+', 'Next', text)
helpers = adapt(source[source.index('Private Sub ParseASCBuffer'):])
start = source.index('        For Each v In addresses')
end = source.index('        Next v', start) + len('        Next v')
output_loop = source[start:end]
output_loop = re.sub(r'(?m)^\s*dst\.Cells\(written \+ 2.*?Value2 = data$', '                CaptureBatch data', output_loop)
output_loop = adapt(output_loop)
cases = (ROOT / 'tools/address_cases.vbs').read_text(encoding='utf-8').strip()
fixture = ('tno=100 > [ 2,100,10,14,0,25,3,4, ] [2,100,10,14,0,25,3,4,]\n'
           '[2,100,11,14,255,1023,15,15,] [SUMMARY] [2,100,10,14,256,0,0,0,]\n'
           'tno=101 > [2,101,10,14,0,25,3,4,] [2,100,12,14,0,0,0,0,]\n'
           'tno=999 > [2,999,10,14,0,0,0,0,]')
fixture_vbs = ' & vbLf & '.join('"' + line + '"' for line in fixture.splitlines())
tests = r'''
Const MAX_PENDING_RECORD = 16384
Const MAX_RESULT_ROWS = 1048575
Dim mTokenRegex, mBitMasks, nums(7), bits(), tnos, addresses, buffer, headerTNO
Dim flagBytes, capacity, sections, skipped, accepted, chunkSize, raw, p, i, count
Dim data(), fields, v, rowIndex, batchRow, b, bitValue, k, j, written
Dim columnCount, batchLimit, batchSize, uniqueCount, captured, t, file, total, maxBuffer
Set mTokenRegex = Nothing
mBitMasks = Array(1,2,4,8,16,32,64,128)
Set tnos = CreateObject("Scripting.Dictionary")
tnos.Add "100", 0: tnos.Add "101", 1
__CASES__
raw = __FIXTURE__
For Each chunkSize In Array(1,7,262144)
 ResetState
 For p = 1 To Len(raw) Step chunkSize
  buffer = buffer & Mid(raw,p,chunkSize): Pump False
 Next
 Pump True
 If addresses.Count<>2 Or accepted<>4 Or skipped<>3 Or sections<>2 Then Die "Counts with chunk size " & chunkSize
 If Not HasFail("10,14,0,25,3,4","100") Or Not HasFail("10,14,0,25,3,4","101") Then Die "Multi-TNO flags"
 If HasFail("11,14,255,1023,15,15","101") Then Die "False positive flag"
 CheckOutput
Next
ResetState
buffer="tno=100 > [2,100": Pump True
If skipped<>1 Then Die "EOF bracket"
ResetState
buffer="tno=100 > [" & String(20000," "): Pump False
If skipped<>1 Or Len(buffer)<>0 Then Die "Pending-record limit"
ResetState
buffer="tno=100": Pump False
buffer=buffer & "0 > [2,1000,10,14,0,0,0,0,]": Pump True
If sections<>0 Or accepted<>0 Then Die "Split unselected TNO"

' Every selectable TNO, including byte boundaries and the last partial byte.
Set tnos = CreateObject("Scripting.Dictionary")
For i=100 To 999: tnos.Add CStr(i), i-100: Next
ResetState
For i=100 To 999
 buffer="tno=" & i & " > [2," & i & ",10,14,0,1,2,3,] [2," & i & ",10,14,0,1,2,3,]"
 Pump False
Next
Pump True
If addresses.Count<>1 Or accepted<>1800 Or flagBytes<>113 Then Die "900 TNO merge"
For i=100 To 999
 If Not HasFail("10,14,0,1,2,3",CStr(i)) Then Die "Missing TNO " & i
Next
CheckOutput

' Force capacity growth and full/partial batch boundaries with 900 TNO columns.
ResetState
For i=0 To 999
 t=100+(i Mod 900)
 buffer="tno=" & t & " > [2," & t & ",10,14,0," & i & ",2,3,]"
 Pump False
Next
Pump True
If addresses.Count<>1000 Or capacity<1000 Then Die "Capacity growth"
For i=0 To 999
 t=i Mod 900
 rowIndex=addresses("10,14,0," & i & ",2,3")
 For b=0 To flagBytes-1
  bitValue=0
  If b=t\8 Then bitValue=mBitMasks(t And 7)
  If bits(b,rowIndex)<>bitValue Then Die "Unexpected bit pattern at address " & i
 Next
Next
CheckOutput
WScript.Echo "PASS: " & count & " address cases; split chunks; 900 TNO bit flags; deduplication; capacity growth; full/partial output batches with blank-vs-1 checks."

If WScript.Arguments.Count>0 Then
 Set tnos=CreateObject("Scripting.Dictionary")
 tnos.Add "100",0: tnos.Add "101",1
 ResetState
 Set file=CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.Arguments(0),1,False,0)
 total=0: maxBuffer=0
 Do Until file.AtEndOfStream
  raw=file.Read(262144): total=total+Len(raw)
  buffer=buffer & raw: Pump False
  If Len(buffer)>maxBuffer Then maxBuffer=Len(buffer)
 Loop
 file.Close: Pump True
 If addresses.Count<>2 Or accepted<>4 Or skipped<>3 Then Die "Large-file result"
 If maxBuffer>16384 Then Die "Unbounded retained buffer"
 CheckOutput
 WScript.Echo "PASS large synthetic file: bytes=" & total & "; max pending chars=" & maxBuffer & "; no log collection or worksheet writes in parser."
End If

Sub ResetState
 Set addresses=CreateObject("Scripting.Dictionary")
 flagBytes=(tnos.Count+7)\8: capacity=2
 ReDim bits(flagBytes-1,capacity-1)
 ' VBA Byte arrays initialize to 0; VBScript Variant arrays initialize Empty,
 ' which is coerced to zero by the same bitwise operations.
 buffer="": headerTNO="": sections=0: accepted=0: skipped=0
End Sub
Sub Pump(eof)
 ParseASCBuffer buffer,eof,headerTNO,sections,accepted,skipped,tnos,addresses,bits,flagBytes,capacity
End Sub
Function HasFail(key,tno)
 Dim col
 col=tnos(tno)
 HasFail=((bits(col\8,addresses(key)) And mBitMasks(col And 7))<>0)
End Function
Sub CheckOutput
 Set captured=CreateObject("Scripting.Dictionary")
 columnCount=6+tnos.Count: uniqueCount=addresses.Count: batchLimit=37
 written=0: batchRow=0: batchSize=batchLimit
 If batchSize>uniqueCount Then batchSize=uniqueCount
 ReDim data(batchSize,columnCount)
 __OUTPUT_LOOP__
 If captured.Count<>uniqueCount Or written<>uniqueCount Then Die "Output row count"
End Sub
Sub CaptureBatch(packet)
 Dim rr,cc,key,tno,expected
 For rr=1 To UBound(packet,1)
  key=""
  For cc=1 To 6
   If cc>1 Then key=key & ","
   key=key & CStr(packet(rr,cc))
  Next
  If Not addresses.Exists(key) Or captured.Exists(key) Then Die "Output address duplicate/missing"
  captured.Add key,True
  For Each tno In tnos
   expected=HasFail(key,tno)
   If expected Then
    If packet(rr,7+tnos(tno))<>1 Then Die "Output lost fail flag"
   Else
    If Not IsEmpty(packet(rr,7+tnos(tno))) Then Die "Non-fail cell must stay blank"
   End If
  Next
 Next
End Sub
Sub Check(bracket,expected)
 Dim actual
 actual=TryParseAddress(bracket,nums): count=count+1
 If actual<>expected Then Die bracket
End Sub
Sub Die(message)
 WScript.Echo "FAIL: " & message: WScript.Quit 1
End Sub
'''.replace('__CASES__', cases).replace('__FIXTURE__', fixture_vbs).replace('__OUTPUT_LOOP__', output_loop)

cli = argparse.ArgumentParser()
cli.add_argument('--large-kb', type=int, default=0)
args = cli.parse_args()
with tempfile.TemporaryDirectory() as directory:
    directory = pathlib.Path(directory)
    script = directory / 'fast_test.vbs'
    script.write_text(tests + '\n' + helpers, encoding='ascii')
    command = ['cscript.exe','//nologo',str(script)]
    if args.large_kb:
        sample=directory/'large.asc'
        padding=('tno=999 >\n'+'padding '*8190+'\n').encode('ascii')
        with sample.open('wb') as output:
            remaining=args.large_kb*1024
            while remaining:
                block=padding[:min(len(padding),remaining)]
                output.write(block)
                remaining-=len(block)
            output.write(b'\n'+fixture.encode('ascii'))
        command.append(str(sample))
    subprocess.run(command,check=True)
