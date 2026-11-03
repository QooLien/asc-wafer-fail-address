from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/BLKVminEngine.bas').read_text()
feed=s[s.index('Public Sub BLK_BeginFile'):s.index('Private Sub ResetParser')]
# Select only BeginFile + Feed; the public interface is earlier in the module.
handler=feed.split('Failed:\n',1)[1].rsplit('End Sub',1)[0]
feed=feed[:feed.index('Failed:')]+'End Sub\n'
feed=feed.replace('On Error GoTo Failed','On Error Resume Next')
lines=[]
for line in feed.splitlines():
 lines.append(line)
 if 'ParseLine ' in line or 'Call FinishRecord' in line:
  lines.append('If Err.Number <> 0 Then\n'+handler+'\nErr.Clear:Exit Sub\nEnd If')
helpers='\n'.join(lines)+'\n'+s[s.index('Private Sub ResetParser'):s.index('Private Function SortedDUTs')]
helpers+=s[s.index('Private Function UnsignedValue'):s.index('Public Sub BLK_ObserveLocations')]
helpers+=s[s.index('Private Function NormalizeASCWhitespace'):s.index('Private Sub SortXY')]
for kind in ('Public','Private'):helpers=helpers.replace(kind+' Sub','Sub').replace(kind+' Function','Function')
helpers=re.sub(r' As (String|Long|Double|Boolean|Object|Variant)\b','',helpers)
helpers=re.sub(r'\b(\w+)\$',r'\1',helpers);helpers=re.sub(r'(\d)#',r'\1',helpers)
helpers=helpers.replace('(0 To ','(').replace('LBound(mSequence)','1')
helpers=re.sub(r'\bNext \w+','Next',helpers)
test='''
Const MAX_EXACT=9007199254740991
Const MAX_LINE=1048576
Dim mTags,mTagHeads,mBlockTag,mSpaces,mDigits,mData,mWanted,mDUTs,mTno,mX,mY,mPath,mLine,mRecordLine
Dim mActive,mRecordKey,mRecordDUT,mValues(),mCount,mCarry,mBLKError,mIgnored,mSequence(2)
Dim phase,b,raw,v,parts,chunk,p,first,sep
Set mTags=Nothing:Set mDigits=Nothing
mSequence(1)="350":mSequence(2)="370"
For Each sep In Array(" ",vbTab,ChrW(&HA0),ChrW(&H3000),ChrW(&H202F))
For Each chunk In Array(1,7,262144)
 For phase=1 To 3
  BLK_BeginFile "WT" & phase
  raw="Column descriptions: Tno= TestNumber X= ChipX Y= ChipY" & vbLf
  raw=raw & "Tno= X= Y=" & vbLf
  raw=raw & "Tno=0 X=1 Y=2 FBC BY BLK=ignored non-FBC metadata" & vbLf
  raw=raw & "Tno=999 X=1 Y=2 FBC BY BLK=not a selected record" & vbLf
  raw=raw & "Tno= 350 DUT= 9 X= 10 Y= 17 FBC BY BLK="
  For b=0 To 255:raw=raw & CStr(phase*1000+b) & " ":Next
  raw=raw & vbLf & "Tno= 350 DUT= 9 X= 10 Y= 17 FBC BY BLK="
  For b=0 To 255:raw=raw & "9 ":Next
  raw=raw & vbLf & "Tno=370 X=-1 Y=2 FBC BY BLK="
  For b=0 To 255:raw=raw & "0 ":Next
  raw=Replace(raw & vbLf," ",sep)
  For p=1 To Len(raw) Step chunk:BLK_Feed Mid(raw,p,chunk),False:Next
  BLK_Feed "",True
  If Len(mBLKError)>0 Then Die mBLKError
  If mData.Count<>2 Or mIgnored<>2 Or mDUTs.Count<>2 Then Die "Selection / ignored TNO0"
  v=mData("350|10|17")
  For b=0 To 255
   If v(b)<>phase*1000+b Then Die "WT / raw index / first duplicate"
  Next
  v=mData("370|-1|2")
  For b=0 To 255:If v(b)<>0 Then Die "Real zero lost"
  Next
 Next
Next
Next
BLK_BeginFile "bad"
BLK_Feed "Tno=350 X=10 Y=10 FBC BY BLK=1 2 3" & vbLf,True
If InStr(mBLKError,"found 3")=0 Then Die "Malformed selected record not reported"
BLK_BeginFile "reset"
If mData.Count<>0 Or Len(mBLKError)>0 Then Die "Cross-file state leaked"
BLK_Feed "Tno=350 X=unknown Y=17 FBC BY BLK=0" & vbLf,True
If InStr(mBLKError,"requires integer X on the same header line")=0 Then Die "Header field diagnostic missing"
BLK_BeginFile "no stale context"
raw="Tno=350 X=10 Y=17 FBC BY BLK="
For b=0 To 255:raw=raw & "0 ":Next
raw=raw & vbLf & "Tno=350 Y=18 FBC BY BLK="
For b=0 To 255:raw=raw & "0 ":Next
BLK_Feed raw & vbLf,True
If InStr(mBLKError,"requires integer X")=0 Or mData.Exists("350|10|18") Then Die "Previous Chip context reused"
BLK_BeginFile "multiline"
raw="Tno= 350 DUT= 9 X= 10 Y= 17 FBC BY BLK="
For b=0 To 255
 raw=raw & CStr(b) & " "
 If b=127 Then raw=raw & vbLf
Next
BLK_Feed raw,True
If Len(mBLKError)>0 Then Die "Multiline / no final newline"
v=mData("350|10|17")
For b=0 To 255:If v(b)<>b Then Die "Multiline raw index"
Next
BLK_BeginFile "257 values"
BLK_Feed raw & "256",True
If InStr(mBLKError,"exceeds 256")=0 Then Die "257 values accepted"
BLK_BeginFile "noninteger"
BLK_Feed "Tno=350 X=10 Y=17 FBC BY BLK=0 1 2." & vbLf,True
If InStr(mBLKError,"Expected non-negative integer")=0 Then Die "Noninteger FBC accepted"
WScript.Echo "PASS single-sheet parser: supplied350/DUT9/X10/Y17 header, five whitespace types, metadata TNO/X/Y ignored, no stale Chip context, multiline256 values, malformed record errors, WT1/2/3 isolation and all256 raw indices."
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
assert 'packet(count, b + 5) = values(b)' in s
assert 'header(1, b + 5) = "BLK " & b' in s
assert 'BuildDUTGroup' not in s and 'JudgeBlock' not in s and 'mGroups' not in s
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'parser.vbs';p.write_text(test+'\n'+helpers,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS single-sheet parser:' in out,out
