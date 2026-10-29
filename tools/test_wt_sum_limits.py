"""Exercise actual write bounds with mock worksheet capacities, not native Excel."""
from pathlib import Path
import re,subprocess,tempfile
root=Path(__file__).resolve().parents[1]
source=(root/'src/SplitWTBySum.bas').read_text()
assert source.index('EnsureModernReport report, tempPath')<source.index('Set statusSheet =')
assert 'WriteWTBlock statusSheet' in source and 'WriteWTBlock dst(p)' in source
helper=source[source.index('Private Sub WriteWTBlock'):source.index('Private Sub SumWTRow')]
helper=helper.replace('Private Sub','Sub')
helper=re.sub(r'ByRef data\(\)','ByRef data',helper)
helper=re.sub(r' As (Worksheet|Long|Variant)\b','',helper)
tests=r'''
Dim ws,data(),saved,writtenRows,writtenColumns,writtenFirst
Set ws=New FakeSheet:ws.Name="WT Status":ws.RowLimit=65536:ws.ColumnLimit=256
ReDim data(553,12)
On Error Resume Next
WriteWTBlock ws,65256,data:saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Legacy row overflow accepted"
ws.RowLimit=1048576:ws.ColumnLimit=16384
WriteWTBlock ws,65256,data
If writtenFirst<>65256 Or writtenRows<>553 Or writtenColumns<>12 Then Die "Modern boundary write differs"
If 65256+writtenRows-1<>65808 Then Die "Reported endpoint mismatch"
On Error Resume Next
WriteWTBlock ws,1048570,data:saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Modern row overflow accepted"
WScript.Echo "PASS actual WriteWTBlock: 65256:65808 rejected for 65536-row sheet; accepted for 1048576-row sheet; true modern overflow rejected."
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
Class FakeCapacity
 Public Count
End Class
Class FakeSheet
 Public RowLimit,ColumnLimit,Name
 Public Property Get Rows
  Dim item:Set item=New FakeCapacity:item.Count=RowLimit:Set Rows=item
 End Property
 Public Property Get Columns
  Dim item:Set item=New FakeCapacity:item.Count=ColumnLimit:Set Columns=item
 End Property
 Public Function Cells(r,c)
  Dim item:Set item=New FakeRange:item.FirstRow=r:Set Cells=item
 End Function
End Class
Class FakeRange
 Public FirstRow,RowCount,ColumnCount
 Public Function Resize(r,c)
  RowCount=r:ColumnCount=c:Set Resize=Me
 End Function
 Public Property Let Value2(value)
  writtenRows=RowCount:writtenColumns=ColumnCount:writtenFirst=FirstRow
  If UBound(value,1)<>RowCount Or UBound(value,2)<>ColumnCount Then Die "Array/range size differs"
 End Property
End Class
'''
with tempfile.TemporaryDirectory() as folder:
    path=Path(folder)/'limits.vbs';path.write_text(tests+'\n'+helper,encoding='ascii')
    result=subprocess.run(['cscript.exe','//nologo',str(path)],capture_output=True,check=True)
    output=(result.stdout+result.stderr).decode(errors='replace');print(output)
    if 'PASS actual WriteWTBlock:' not in output:raise RuntimeError('Capacity test did not complete')
assert 131072//(6+3*77)==553
assert 2+118*553==65256
print('PASS: 77 TNOs/WT reproduce original batch start 65256 and crossing of legacy row limit.')
