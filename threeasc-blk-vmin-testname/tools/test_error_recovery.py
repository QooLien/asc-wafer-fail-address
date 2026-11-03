"""Exercise the shipped FBC failure handler, including its state mutations."""
from pathlib import Path
import subprocess,tempfile
root=Path(__file__).resolve().parents[1]
s=(root/'src/BLKVminEngine.bas').read_text()
feed=s[s.index('Public Sub BLK_Feed'):s.index('Public Function BLK_WritePhase')]
handler=feed.split('Failed:\n',1)[1].rsplit('End Sub',1)[0]
release=s[s.index('Public Sub BLK_Release'):s.index('Private Sub ReadConfiguration')]
assert 'Set mGroups = Nothing' not in feed
assert 'Set mGroups = Nothing' in release
code='''
Dim mGroups,mSheetNames,mNextRows,mBLKError,mLine,mCarry,mActive,k,n
Set mGroups=CreateObject("Scripting.Dictionary"):mGroups.Add "W-Low pBL",True
Set mSheetNames=CreateObject("Scripting.Dictionary"):mSheetNames.Add "W-Low pBL","BLK_Distribution(W-Low pBL)"
Set mNextRows=CreateObject("Scripting.Dictionary"):mNextRows.Add "W-Low pBL",2
mLine=123:mCarry="partial":mActive=True
TriggerParseError
If InStr(mBLKError,"Expected 256 FBC values")=0 Or InStr(mBLKError,"line=123")=0 Then Die "Original diagnostic lost"
If mCarry<>"" Or mActive Then Die "Record state not reset"
n=0:For Each k In mGroups:n=n+1:Next
If n<>1 Or mSheetNames("W-Low pBL")<>"BLK_Distribution(W-Low pBL)" Or mNextRows("W-Low pBL")<>2 Then Die "Output state destroyed"
WScript.Echo "PASS original FBC failure handler: parse error and line retained; partial record reset; group/sheet/row dictionaries survive and remain iterable without 424."
Sub TriggerParseError
 On Error Resume Next
 Err.Raise vbObjectError+312,,"Expected 256 FBC values"
'''+handler+'''
End Sub
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'recovery.vbs';p.write_text(code,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS original FBC failure handler:' in out,out
