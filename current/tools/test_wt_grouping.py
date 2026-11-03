from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/BLKVminEngine.bas').read_text()
final=s[s.index('Public Function BLK_Finalize'):s.index('Private Function BuildTNORanks')]
keys=re.findall(r'\.SortFields.Add Key:=ws.Cells\(2, (\d+)\)',final)
assert keys==['3','4','261','1'],keys
assert '.SetRange ws.Cells(1, 1).Resize(lastRow, 261)' in final
assert 'ws.Cells(1, 261).Resize(lastRow, 1).ClearContents' in final
assert 'ws.Cells(1, 1).Resize(lastRow, 260).AutoFilter' in final
helper=s[s.index('Private Function BuildTNORanks'):s.index('Public Sub BLK_Release')]
helper=helper.replace('Private Function','Function')
helper=re.sub(r' As (Variant|Long|String|Object)\b','',helper)
helper=helper.replace('(1 To ','(').replace(', 1 To ',', ')
helper=re.sub(r'\bNext \w+','Next',helper)
code='''
Dim ranks,values(4,1),result,i,saved
Set ranks=CreateObject("Scripting.Dictionary")
ranks.Add "370",1:ranks.Add "350",2:ranks.Add "351",3
values(1,1)=350:values(2,1)=370:values(3,1)=351:values(4,1)=350
result=BuildTNORanks(values,ranks,4)
If result(1,1)<>2 Or result(2,1)<>1 Or result(3,1)<>3 Or result(4,1)<>2 Then Die "Config sequence changed"
result=BuildTNORanks(350,ranks,1)
If result(1,1)<>2 Then Die "Single-row Value2 scalar"
On Error Resume Next
result=BuildTNORanks(999,ranks,1):saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Unconfigured TNO accepted"
WScript.Echo "PASS WT grouping ranks: original Config order, repeated TNO rows, one-row scalar and invalid TNO rejection."
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'group.vbs';p.write_text(code+'\n'+helper,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS WT grouping ranks:' in out,out
# Independent sort-key oracle: every full 260-field row stays intact.
seq={370:1,350:2,351:3}
rows=[(wt,tno,x,y,*[wt*1000000+tno*1000+b for b in range(256)]) for wt in [3,1,2] for x,y in [(11,17),(10,17)] for tno in [350,351,370]]
original=set(rows)
rows.sort(key=lambda a:(a[2],a[3],seq[a[1]],a[0]))
assert set(rows)==original
for i in range(0,len(rows),3):
 group=rows[i:i+3]
 assert [a[0] for a in group]==[1,2,3]
 assert len({a[1:4] for a in group})==1
 assert all(a[4]==a[0]*1000000+a[1]*1000 and a[-1]==a[0]*1000000+a[1]*1000+255 for a in group)
assert rows[0][1]==370
print('PASS sort configuration: Chip X/Y, Config TNO rank, WT1/2/3; entire256-FBC rows preserved; temporary helper removed. Native sort is covered by test_native_maps.ps1.')
