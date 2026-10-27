"""Exercise actual metadata/name VBA helpers in VBScript; no Excel SaveAs call."""
from pathlib import Path
import re
import subprocess
import tempfile

root = Path(__file__).resolve().parents[1]
source = (root / 'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
block = source[source.index('Private Sub CaptureASCMetadata'):source.index('Private Sub ParseASCBuffer')]
block += '\n' + source[source.index('Private Function CleanSpace'):]
block = block.replace('Private Function','Function').replace('Private Sub','Sub')
block = re.sub(r' As (Object|String|Variant|Long|Double|Boolean)\b','',block)
block = re.sub(r'\b(\w+)\$',r'\1',block)
block = re.sub(r'(\d)#',r'\1',block)
block = re.sub(r'\bNext \w+','Next',block)
tests = r'''
Const MAX_PENDING_RECORD=16384
Dim mHeaderRegex, metadata, text, raw, size, p, result, fso, folder, target, file, saved, n
Set mHeaderRegex=Nothing
raw="LOT= L12345 WNO= 3 STEP = WTTAAAA Program = " & ChrW(20081) & ChrW(30908) & "VER_FB" & vbCrLf & _
    "LOT= SECOND WNO= 99 STEP=WTTBBBB Program=VER_ZZ" & vbCrLf
For Each size In Array(1,7,262144)
 Set metadata=CreateObject("Scripting.Dictionary"): text=""
 For p=1 To Len(raw) Step size
  text=text & Mid(raw,p,size)
  CaptureASCMetadata text,False,metadata
 Next
 CaptureASCMetadata text,True,metadata
 result=BuildResultBaseName(metadata)
 If result<>"L12345_#03_AAAA_Ver_FB" Then Die "First occurrence / WTT prefix / noise / split chunks: " & result
Next
Set metadata=CreateObject("Scripting.Dictionary")
text="Program=garbledVER_RC2_xyz" & vbLf & "STEP=WTTABCD" & vbLf & "WNO=10" & vbLf & "LOT=LOT-01"
CaptureASCMetadata text,True,metadata
If BuildResultBaseName(metadata)<>"LOT-01_#10_ABCD_Ver_RC2_xyz" Then Die "Program order or EOF"
metadata("WNO")="0"
If InStr(BuildResultBaseName(metadata),"_#00_")=0 Then Die "WNO zero"
metadata("WNO")="09"
If InStr(BuildResultBaseName(metadata),"_#09_")=0 Then Die "Already padded"
metadata("LOT")="A/B:C"
If Left(BuildResultBaseName(metadata),5)<>"A_B_C" Then Die "Invalid filename characters"
metadata("LOT")="A[B]"
If Left(BuildResultBaseName(metadata),4)<>"A_B_" Then Die "Excel bracket filename characters"
metadata("WNO")="invalid"
On Error Resume Next
result=BuildResultBaseName(metadata): saved=Err.Number: Err.Clear
On Error GoTo 0
If saved=0 Then Die "Invalid first WNO accepted"
metadata("WNO")="3": metadata("PROGRAM")="no version"
On Error Resume Next
result=BuildResultBaseName(metadata): saved=Err.Number: Err.Clear
On Error GoTo 0
If saved=0 Then Die "Missing version accepted"
metadata.Remove "LOT"
On Error Resume Next
result=BuildResultBaseName(metadata): saved=Err.Number: Err.Clear
On Error GoTo 0
If saved=0 Then Die "Missing first header accepted"
Set fso=CreateObject("Scripting.FileSystemObject")
folder=WScript.Arguments(0)
target=UniqueResultPath(fso,folder,"L12345_#03_AAAA_Ver_FB")
Set file=fso.CreateTextFile(target,False): file.Close
If fso.GetFileName(UniqueResultPath(fso,folder,"L12345_#03_AAAA_Ver_FB"))<>"L12345_#03_AAAA_Ver_FB_001.xlsm" Then Die "Collision handling"
For n=1 To 999
 Set file=fso.CreateTextFile(fso.BuildPath(folder,"L12345_#03_AAAA_Ver_FB_" & Right("000" & n,3) & ".xlsm"),False): file.Close
Next
If fso.GetFileName(UniqueResultPath(fso,folder,"L12345_#03_AAAA_Ver_FB"))<>"L12345_#03_AAAA_Ver_FB_1000.xlsm" Then Die "Collision sequence 1000"
WScript.Echo "PASS: first headers, chunk boundaries, noisy Program versions, WTT removal, WNO padding, missing/invalid headers, safe names and non-overwrite paths."
Sub Die(message)
 WScript.Echo "FAIL: " & message: WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as directory:
    script=Path(directory)/'metadata_test.vbs'
    script.write_text(tests+'\n'+block,encoding='ascii')
    subprocess.run(['cscript.exe','//nologo',str(script),directory],check=True)
