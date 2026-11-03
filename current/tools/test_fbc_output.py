from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/BLKVminEngine.bas').read_text()
a=s.index('    Set duts = SortedDUTs(mDUTs)');b=s.index('    Select Case phase',a)
body=s[a:b].replace('Set duts = SortedDUTs(mDUTs)','Set duts = mDUTs')
body=body.replace('(1 To ','(').replace(', 1 To ',', ').replace('LBound(mSequence)','1')
body=re.sub(r'\bNext \w+','Next',body)
code='''
Class SinkRange
 Public owner,firstRow,take
 Public Function Resize(nr,nc)
  take=nr:Set Resize=Me
 End Function
 Public Property Let Value2(packet)
  owner.Check packet,firstRow,take
 End Property
End Class
Class SinkSheet
 Public total,phase
 Public Function Cells(r,c)
  Dim cell:Set cell=New SinkRange:Set cell.owner=Me:cell.firstRow=r:Set Cells=cell
 End Function
 Public Sub Check(packet,firstRow,take)
  Dim r,b,tno,dutIndex,expected
  For r=1 To take
   tno=100+((total Mod 1026) Mod 513)
   dutIndex=Int((total Mod 1026)/513)
   If packet(r,1)<>"WT" & phase Or packet(r,2)<>tno Or packet(r,3)<>dutIndex Or packet(r,4)<>10 Then Die "WT/TNO/XY mapping"
   For b=0 To 255
    expected=phase*1000000+dutIndex*600000+tno*1000+b
    If packet(r,b+5)<>expected Then Die "BLK mapping / packet boundary"
   Next
   total=total+1
  Next
 End Sub
End Class
Dim ws,mDUTs,mData,mSequence(513),mNextRow,phase,dutIndex,i,b,v(255),dut,xy,values,key
Dim duts,startRow,remaining,take,packet(),count
Set ws=New SinkSheet:ws.total=0:mNextRow=2
For i=1 To 513:mSequence(i)=CStr(99+i):Next
For phase=1 To 3
 ws.phase=phase
 Set mDUTs=CreateObject("Scripting.Dictionary"):Set mData=CreateObject("Scripting.Dictionary")
 For dutIndex=0 To 1
  mDUTs.Add CStr(dutIndex) & "|10",True
  For i=1 To 513
   For b=0 To 255:v(b)=phase*1000000+dutIndex*600000+CLng(mSequence(i))*1000+b:Next
   mData.Add mSequence(i) & "|" & dutIndex & "|10",v
  Next
 Next
 count=0
'''+body+'''
Next
If ws.total<>3078 Or mNextRow<>3080 Then Die "Row count"
WScript.Echo "PASS actual output loop: 3078 WT/TNO/Chip rows, all256 BLK fields, WT transitions, 512-row batches and final partial batch."
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'output.vbs';p.write_text(code,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS actual output loop:' in out,out
