from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
s=(root/'src/BLKVminEngine.bas').read_text()
helper=s[s.index('Private Function BuildComparisonGroup'):s.index('Public Function BLK_AddressTNOs')]
helper=helper.replace('Optional ByVal addVmin As Boolean = True','ByVal addVmin')
helper=helper.replace('Private Function','Function')
helper=re.sub(r'ByRef (tnos|volts)\(\)',r'ByRef \1',helper)
helper=re.sub(r' As (Variant|String|Long|Double|Object)\b','',helper)
helper=helper.replace('(1 To ','(').replace(', 1 To ',', ')
helper=re.sub(r'\bNext \w+','Next',helper)
code='''
Dim source(12,260), rows, tnos(2), volts(2), p,i,b,k,result,r
Set rows=CreateObject("Scripting.Dictionary")
tnos(1)="350":tnos(2)="351":volts(1)=1.2:volts(2)=1.035
For p=1 To 3
 For i=1 To 3
  r=(p-1)*3+i
  If i<3 Then k="10|10|" & tnos(i) & "|" & p Else k="10|10|Vmin|" & p
  rows.Add k,r
  For b=0 To 255:source(r,b+5)=p*100000+i*1000+b:Next
 Next
Next
rows.Remove "10|10|351|2"
result=BuildComparisonGroup(source,rows,"10|10","W-Low pBL",tnos,volts,True)
If UBound(result,1)<>9 Or UBound(result,2)<>262 Then Die "Shape"
For i=1 To 3
 For p=1 To 3
  r=(i-1)*3+p
  If result(r,2)<>"W-Low pBL" Or result(r,4)<>10 Or result(r,5)<>10 Or result(r,6)<>"WT" & p Then Die "Metadata"
  If i<3 Then
   If CStr(result(r,1))<>tnos(i) Or result(r,3)<>volts(i) Then Die "TNO voltage"
  Else
   If result(r,1)<>"X10Y10 Vmin" Then Die "Vmin label"
  End If
  For b=0 To 255
   If i=2 And p=2 Then
    If result(r,b+7)<>"Not Found" Then Die "Missing treated as zero"
   Else
    If result(r,b+7)<>p*100000+i*1000+b Then Die "WT/TNO/BLK mismatch"
   End If
  Next
 Next
Next
result=BuildComparisonGroup(source,rows,"-1|2","R-Low pBL",tnos,volts,True)
If result(1,7)<>"Not Found" Or result(9,262)<>"No Data" Then Die "Absent DUT"
result=BuildComparisonGroup(source,rows,"10|10","W-Low pBL",tnos,volts,False)
If UBound(result,1)<>6 Then Die "Function unexpectedly has Vmin rows"
WScript.Echo "PASS comparison layout: all256 blocks, three WT adjacent per TNO, voltage/Func/XY identity, independent Vmin values, missing WT and DUT."
Function IIf(test,a,b)
 If test Then IIf=a Else IIf=b
End Function
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'comparison.vbs';p.write_text(code+'\n'+helper,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 output=(result.stdout+result.stderr).decode('cp950',errors='replace');print(output.encode('ascii','backslashreplace').decode())
 assert 'PASS comparison layout:' in output,output
