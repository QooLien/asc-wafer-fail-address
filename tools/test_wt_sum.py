"""Test actual independent WT Sum and per-WT output helpers against fixture truth."""
from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
source=(root/'src/SplitWTBySum.bas').read_text()
helpers=source[source.index('Private Sub SumWTRow'):source.index('Private Sub FormatSumSheet')]
helpers+=source[source.index('Private Sub ValidateWTHeaders'):]
def adapt(text):
    text=text.replace('Private Sub','Sub').replace('Private Function','Function')
    text=re.sub(r'ByRef (colWT|sums|colMap|output)\(\)',r'ByRef \1',text)
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object|Byte)\b','',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text)
    text=re.sub(r'(\d)#',r'\1',text)
    text=text.replace('(1 To ','(').replace(', 1 To ',', ')
    text=re.sub(r'\bNext \w+','Next',text)
    return text.replace('IsError(', 'FixtureCellError(')
tests=r'''
Dim n,lastCol,header(),colWT(),colMap(),packet(),sums(),truth(),output(),counts(3)
Dim p,i,c,r,phase,mask,flag,expected,rows,key,sourceRow,flags,junk,saved
For Each n In Array(1,3,9,900)
 lastCol=6+3*n:ReDim header(1,lastCol)
 junk=Array("X","Y","BLK","WL","BL","IO")
 For c=1 To 6:header(1,c)=junk(c-1):Next
 For p=1 To 3
  For i=0 To n-1:header(1,7+(p-1)*n+i)="WT" & (4-p) & vbLf & (100+((i*137) Mod 900)):Next
 Next
 ValidateWTHeaders header,colWT
 ReDim colMap(3,n)
 For p=1 To 3:counts(p)=0:Next
 For c=7 To lastCol
  p=colWT(c):counts(p)=counts(p)+1:colMap(p,counts(p))=c
 Next
 ReDim packet(8,lastCol):ReDim sums(8,3):ReDim truth(8,3)
 For r=1 To 8
  packet(r,1)=10:packet(r,2)=14:packet(r,3)=0:packet(r,4)=r:packet(r,5)=2:packet(r,6)=3
  mask=r-1
  For p=1 To 3
   For i=1 To n
    If (mask And (2^(p-1)))<>0 And (i=1 Or ((i+r+p) Mod 3)=0) Then
     packet(r,colMap(p,i))=CDbl(1)
     truth(r,p)=truth(r,p)+1
    End If
   Next
  Next
  SumWTRow packet,r,colWT,sums
  For p=1 To 3
   If sums(r,p)<>truth(r,p) Then Die "Wrong independent WT Sum"
  Next
 Next
 For p=1 To 3
  expected=0
  For r=1 To 8:If truth(r,p)>0 Then expected=expected+1
  Next
  ReDim output(expected,n+8)
  BuildWTFailBatch packet,sums,colMap,p,n,output,rows
  If rows<>expected Then Die "Wrong Fail sheet count"
  For r=1 To rows
   sourceRow=CLng(output(r,4))
   If truth(sourceRow,p)<=0 Then Die "Pass address included in Fail sheet"
   For c=1 To 6
    If CStr(output(r,c))<>CStr(packet(sourceRow,c)) Then Die "Address modified"
   Next
   flags=0
   For i=1 To n
    If CStr(output(r,6+i))<>CStr(packet(sourceRow,colMap(p,i))) Then Die "Wrong TNO column value"
    If Not IsEmpty(output(r,6+i)) Then flags=flags+1
   Next
   If output(r,n+7)<>truth(sourceRow,p) Or flags<>output(r,n+7) Or output(r,n+8)<>"Fail" Then Die "Wrong Fail Sum / status"
  Next
  If sums(8,p)=0 Then Die "All-WT fail address wrongly excluded"
 Next
 WScript.Echo "PASS independent WT sums: " & n & " TNOs; reversed WT columns; all 8 combinations; multi-TNO sums; overlapping Fail sheets; exact per-TNO detail and address copies."
Next
ReDim packet(1,9):ReDim sums(1,3):ReDim colWT(9)
colWT(7)=1:colWT(8)=2:colWT(9)=3
packet(1,7)=" 1 ":packet(1,8)=" "
SumWTRow packet,1,colWT,sums
If sums(1,1)<>1 Or sums(1,2)<>0 Or sums(1,3)<>0 Then Die "Text/blank handling"
For Each junk In Array(2,True,False,0,1.5)
 packet(1,7)=junk
 On Error Resume Next
 SumWTRow packet,1,colWT,sums:saved=Err.Number:Err.Clear
 On Error GoTo 0
 If saved=0 Then Die "Invalid flag accepted"
Next
WScript.Echo "PASS WT Sum validation: text/blank handling and invalid flag rejection."
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
Function FixtureCellError(value)
 FixtureCellError=False
End Function
'''
with tempfile.TemporaryDirectory() as folder:
    path=Path(folder)/'sum.vbs';path.write_text(adapt(tests+'\n'+helpers),encoding='ascii')
    result=subprocess.run(['cscript.exe','//nologo',str(path)],capture_output=True,check=True)
    output=(result.stdout+result.stderr).decode(errors='replace');print(output)
    if 'PASS WT Sum validation:' not in output:raise RuntimeError('WT Sum test did not complete')
