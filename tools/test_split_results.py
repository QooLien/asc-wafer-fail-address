"""Execute actual stage-two header/classifier/copy code in VBScript."""
from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
source=(root/'src/SplitThreeASCResults.bas').read_text()
helpers=source[source.index('Private Sub ValidateWTHeaders'):source.index('Private Sub FormatWTOutput')]
a=source.index('                rr = 0')
copy=source[a:source.index('                targets(g).Cells',a)]
def adapt(text):
    text=text.replace('Private Sub','Sub').replace('Private Function','Function')
    # VBScript has no Excel/VBA IsError; fixtures contain no Excel error variants.
    text=text.replace('IsError(', 'FixtureCellError(')
    text=re.sub(r'ByRef colWT\(\)', 'ByRef colWT',text)
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object|Byte)\b','',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text)
    text=re.sub(r'(\d)#',r'\1',text)
    text=text.replace('(1 To ','(').replace(', 1 To ',', ')
    return re.sub(r'\bNext \w+','Next',text)
tests=r'''
Dim header(),packet(),colWT(),classes(),output(),seen,key,mode,n,c,p,i,r,g,rr,count,lastCol
Dim expected,mask,t,actual,counts(4),saved,dummy,expectedPhase
For Each n In Array(1,3,9,900)
 lastCol=6+3*n
 ReDim header(1,lastCol)
 dummy=Array("X","Y","BLK","WL","BL","IO")
 For c=1 To 6:header(1,c)=dummy(c-1):Next
 ' Permute TNO order and reverse the WT column groups to verify header-based mapping.
 For p=1 To 3
  For i=0 To n-1
   header(1,7+(p-1)*n+i)="WT" & (4-p) & vbLf & (100+((i*137) Mod 900))
  Next
 Next
 ValidateWTHeaders header,colWT
 count=8:ReDim packet(count,lastCol):ReDim classes(count)
 For r=1 To count
  packet(r,1)=10:packet(r,2)=14:packet(r,3)=0:packet(r,4)=r:packet(r,5)=2:packet(r,6)=3
  mask=r-1
  For p=1 To 3
   If (mask And (2^(p-1)))<>0 Then
    For c=7 To lastCol
     If colWT(c)=p Then packet(r,c)=1:Exit For
    Next
   End If
  Next
  actual=ClassifyWTRow(packet,r,colWT)
  Select Case mask
   Case 0:expected=0
   Case 1:expected=1
   Case 2:expected=2
   Case 4:expected=3
   Case Else:expected=4
  End Select
  If actual<>expected Then Die "Wrong WT combination class"
  If actual=0 Then actual=4
  classes(r)=actual:counts(actual)=counts(actual)+1
 Next
 Set seen=CreateObject("Scripting.Dictionary")
 For g=1 To 4
  ReDim output(count,lastCol+1)
__COPY__
  For r=1 To rr
   key=AddressKey(output,r)
   If seen.Exists(key) Then Die "Address duplicated across groups"
   seen.Add key,True
   i=CLng(output(r,4))
   If classes(i)<>g Then Die "Row placed on wrong sheet"
   For c=1 To lastCol
    If CStr(output(r,c))<>CStr(packet(i,c)) Then Die "Original cell moved or modified"
   Next
   If g=3 Then
    For c=7 To lastCol
     If colWT(c)<>3 And Not IsEmpty(output(r,c)) Then Die "WT3 sheet has earlier WT fail"
    Next
   End If
   If g=4 Then
    If Len(output(r,lastCol+1))=0 Then Die "Conflict reason missing"
   End If
  Next
 Next
 If seen.Count<>count Then Die "Reconciled row count differs"
 WScript.Echo "PASS stage-two split: " & n & " TNOs; reversed WT columns; all eight combinations; original cells preserved; no lost or duplicated addresses."
Next
' Validate guardrails using final generated header.
header(1,8)=header(1,7)
On Error Resume Next
ValidateWTHeaders header,colWT:saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Duplicate WT/TNO header accepted"
ReDim packet(1,9):ReDim colWT(9)
colWT(7)=1:colWT(8)=2:colWT(9)=3
packet(1,7)=2
On Error Resume Next
actual=ClassifyWTRow(packet,1,colWT):saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Invalid fail value accepted"
WScript.Echo "PASS header duplicate and non-blank/non-1 rejection."
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
Function FixtureCellError(value)
 FixtureCellError=False
End Function
'''
with tempfile.TemporaryDirectory() as folder:
    path=Path(folder)/'split.vbs'
    path.write_text(adapt(tests.replace('__COPY__',copy)+'\n'+helpers),encoding='ascii')
    result=subprocess.run(['cscript.exe','//nologo',str(path)],capture_output=True,check=True)
    output=(result.stdout+result.stderr).decode(errors='replace')
    print(output)
    if 'PASS header duplicate and non-blank/non-1 rejection.' not in output:
        raise RuntimeError('Stage-two tests did not complete')
