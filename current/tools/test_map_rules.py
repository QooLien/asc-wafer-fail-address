from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/BLKVminEngine.bas').read_text()
def function(name):
 a=s.index('Private Function '+name);return s[a:s.index('End Function',a)+len('End Function')]
helpers='\n'.join(function(n) for n in ['ReadStopPatterns','CountFBCRow','MapResult','MapStatus','MapColor','FailMapSheetName','UnsignedValue','Coordinate'])
helpers=helpers.replace('Private Function','Function')
helpers=re.sub(r' As (Variant|String|Long|Double|Object|Boolean|Worksheet)\b','',helpers)
helpers=re.sub(r'\b(\w+)\$',r'\1',helpers);helpers=re.sub(r'(\d)#',r'\1',helpers)
helpers=re.sub(r'\bNext \w+','Next',helpers)
code='''
Class ConfigRange
 Public data
 Public Property Get Value2
  Value2=data
 End Property
End Class
Class FakeConfig
 Public data
 Public Function Range(address)
  If address<>"D2:G901" Then Die "Monitor columns read for judgment"
  Dim cell:Set cell=New ConfigRange:cell.data=data:Set Range=cell
 End Function
End Class
Dim cfg,settings(3,4),patterns,wanted,counts,errors,data(1,260),result,b,saved,mDigits,used,a,c
Set mDigits=Nothing:Set cfg=New FakeConfig
settings(1,1)=500:settings(1,2)="Scan":settings(1,3)=1.2
settings(2,1)=501:settings(2,2)="Scan":settings(2,3)=1.1:settings(2,4)="Scan"
settings(3,1)=600:settings(3,2)="SRAM":settings(3,3)=1.2:settings(3,4)="SRAM"
cfg.data=settings:Set patterns=ReadStopPatterns(cfg,wanted)
If patterns.Count<>2 Or wanted.Exists("500") Or Not wanted.Exists("501") Then Die "Wrong Stop endpoint"
a=patterns("Scan"):If a(0)<>"501" Then Die "500 selected instead of501"
For b=5 To 260:data(1,b)=0:Next
If CountFBCRow(data,1)<>0 Then Die "Zero must Pass"
data(1,5)=1:data(1,22)=3:data(1,260)=999999
If CountFBCRow(data,1)<>3 Then Die "FBC magnitude summed instead of BLK count"
Set counts=CreateObject("Scripting.Dictionary"):Set errors=CreateObject("Scripting.Dictionary")
counts.Add "WT1|500|10|17",256:counts.Add "WT1|501|10|17",0
counts.Add "WT2|501|10|17",3
result=MapResult(counts,errors,"WT1","501","10|17")
If result<>0 Or MapStatus(result)<>"Pass" Then Die "Prior500 fail incorrectly accumulated"
result=MapResult(counts,errors,"WT2","501","10|17")
If result<>3 Or MapStatus(result)<>"Fail" Then Die "Endpoint count"
result=MapResult(counts,errors,"WT3","501","10|17")
If result<>"No Data" Or MapColor(result)<>RGB(191,191,191) Then Die "Missing WT must be gray"
counts.Add "WT3|500|10|17",100
If MapResult(counts,errors,"WT3","501","10|17")<>"No Data" Then Die "Other TNO cannot supply Stop result"
errors.Add "WT2",True
If MapResult(counts,errors,"WT2","501","10|17")<>"Data Error" Then Die "Parse failure mislabeled as untested"
data(1,5)=""
On Error Resume Next
result=CountFBCRow(data,1):saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Blank FBC treated as zero"
settings(3,4)="scan":cfg.data=settings
On Error Resume Next
Set result=ReadStopPatterns(cfg,wanted):saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Duplicate Pattern Name accepted"
Set used=CreateObject("Scripting.Dictionary"):used.CompareMode=vbTextCompare
a=FailMapSheetName("Scan",used):If a<>"Scan Fail Map" Then Die "Map title"
a=FailMapSheetName("ABCDEFGHIJKLMNOPQRSTUVWXYZ1",used):c=FailMapSheetName("ABCDEFGHIJKLMNOPQRSTUVWXYZ2",used)
If Len(a)>31 Or Len(c)>31 Or a=c Then Die "Sheet name collision"
WScript.Echo "PASS Stop Sign Maps: Function D:G only; named endpoint501, no500 accumulation; count0/3; WT3 no data gray; missing endpoint stays missing; parse failure yellow; invalid FBC and duplicate Pattern rejected; safe Map names."
Function IsError(value)
 IsError=(VarType(value)=10)
End Function
Function RGB(r,g,b)
 RGB=r+g*256+b*65536
End Function
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'maps.vbs';p.write_text(code+'\n'+helpers,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS Stop Sign Maps:' in out,out
