from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1];s=(root/'src/BLKVminEngine.bas').read_text()
helper=s[s.index('Private Function BuildDUTGroup'):s.index('Private Function SortedDUTs')]
helper+=s[s.index('Private Function HasMonitor'):]
helper=helper.replace('Private Function','Function')
helper=re.sub(r'ByRef (tnos|volts|cache|present)\(\)',r'ByRef \1',helper)
helper=re.sub(r' As (Variant|String|Long|Double|Boolean|Object)\b','',helper)
helper=re.sub(r'\b(\w+)\$',r'\1',helper)
helper=helper.replace('(1 To ','(').replace(', 1 To ',', ')
for name in ('tnos','present'):helper=helper.replace(f'LBound({name})','1')
helper=re.sub(r'\bNext \w+','Next',helper)
tests='''
Dim data,monitor,tnos(3),volts(3),v(255),i,b,result
Set data=CreateObject("Scripting.Dictionary"):Set monitor=CreateObject("Scripting.Dictionary")
tnos(1)="350":tnos(2)="900":tnos(3)="351"
volts(1)=1.2:volts(2)=9.9:volts(3)=1.035
monitor.Add "350",True:monitor.Add "351",True
For i=1 To 3
 For b=0 To 255:v(b)=0:Next
 If i=2 Then
  For b=0 To 255:v(b)=1000:Next
 End If
 If i=3 Then v(0)=5:v(255)=5
 data.Add tnos(i) & "|10|10",v
Next
result=BuildItemDUTGroup(data,tnos,volts,"10|10",5,monitor)
If UBound(result,1)<>4 Then Die "Monitor Vmin missing"
If result(4,5)<>1.2 Or result(4,260)<>1.2 Or result(4,6)<>"Pass" Then Die "Function contaminates threshold or previous voltage"
If result(2,5)<>1000 Or result(2,260)<>1000 Then Die "Function FBC lost"
result=BuildItemDUTGroup(data,tnos,volts,"10|10",0,monitor)
If result(4,5)<>1.2 Or result(4,260)<>1.2 Then Die "Function contaminates Legacy"
Set monitor=CreateObject("Scripting.Dictionary")
result=BuildItemDUTGroup(data,tnos,volts,"10|10",0,monitor)
If UBound(result,1)<>3 Or HasMonitor(tnos,monitor) Then Die "Function-only group judged"
monitor.Add "350",True:monitor.Add "351",True
data.Remove "350|10|10"
result=BuildItemDUTGroup(data,tnos,volts,"10|10",5,monitor)
If result(4,5)<>"Missing Data" Or result(1,5)<>"Not Found" Then Die "Missing Monitor treated as zero"
WScript.Echo "PASS dual item groups: Monitor-only Legacy/threshold/previous voltage; Function FBC retained, never judged; missing data preserved; BLK0/255 correct."
Function IIf(test,a,b)
 If test Then IIf=a Else IIf=b
End Function
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'items.vbs';p.write_text(tests+'\n'+helper,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 output=(result.stdout+result.stderr).decode('cp950',errors='replace');print(output.encode('ascii','backslashreplace').decode())
 assert 'PASS dual item groups:' in output,output
