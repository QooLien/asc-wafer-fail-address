from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1];s=(root/'src/BLKVminEngine.bas').read_text()
helper=s[s.index('Public Function BLK_AddressTNOs'):s.index('Private Function HasMonitor')]
helper+=s[s.index('Private Function UnsignedValue'):s.index('Private Function Coordinate')]
helper=helper.replace('Private Sub','Sub').replace('Private Function','Function').replace('Public Function','Function')
helper=re.sub(r'ByRef (tnos|names|volts)\(\)',r'ByRef \1',helper)
helper=re.sub(r' As (Variant|String|Long|Double|Boolean|Object)\b','',helper)
helper=re.sub(r'\b(\w+)\$',r'\1',helper)
helper=helper.replace('LBound(mSequence)','1')
helper=re.sub(r'\bNext \w+','Next',helper)
tests='''
Dim mSequence,mMonitorTNOs,mDigits,seen,tnos(),names(900),volts(900),count,a(3,3),f(1,3),indices,saved
ReDim tnos(900)
Set mMonitorTNOs=CreateObject("Scripting.Dictionary"):Set seen=CreateObject("Scripting.Dictionary"):Set mDigits=Nothing
a(1,1)=350:a(1,2)="W-Low pBL":a(1,3)=1.2
a(3,1)=351:a(3,2)="W-Low pBL":a(3,3)=1.035
f(1,1)=900:f(1,2)="Function":f(1,3)=1.2
ReadItemColumns a,True,seen,tnos,names,volts,count
ReadItemColumns f,False,seen,tnos,names,volts,count
If count<>3 Or tnos(1)<>"350" Or tnos(2)<>"351" Or tnos(3)<>"900" Then Die "Sequence"
If mMonitorTNOs.Count<>2 Or mMonitorTNOs.Exists("900") Then Die "Monitor membership"
ReDim Preserve tnos(3):mSequence=tnos:Set indices=BLK_AddressTNOs()
If indices("350")<>0 Or indices("351")<>1 Or indices("900")<>2 Then Die "TNO/address index mismatch"
On Error Resume Next
ReadItemColumns f,False,seen,tnos,names,volts,count
saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Duplicate accepted"
f(1,1)=999:f(1,2)="":f(1,3)=1.2
On Error Resume Next
ReadItemColumns f,False,seen,tnos,names,volts,count
saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Missing Func accepted"
f(1,1)="100.5":f(1,2)="Function"
On Error Resume Next
ReadItemColumns f,False,seen,tnos,names,volts,count
saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Noninteger accepted"
WScript.Echo "PASS dual Config: sparse lists, Monitor then Function order, zero-based address indices, monitor membership, duplicate and invalid input rejection."
Function IIf(test,a,b)
 If test Then IIf=a Else IIf=b
End Function
Function IsError(value)
 IsError=(VarType(value)=10)
End Function
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'config.vbs';p.write_text(tests+'\n'+helper,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 output=(result.stdout+result.stderr).decode('cp950',errors='replace');print(output.encode('ascii','backslashreplace').decode())
 assert 'PASS dual Config:' in output,output
