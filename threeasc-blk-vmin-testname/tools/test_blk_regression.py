"""Exercise actual VBA parser and output/judge code in VBScript, with Python oracle."""
from pathlib import Path
import re,subprocess,tempfile,json
ROOT=Path(__file__).resolve().parents[1]
source=(ROOT/'src/BLKVminEngine.bas').read_text()
seq=[103,101,104,102];volts=[.60,.65,.70,.75]
records={}
for pos,tno in enumerate(seq):
    a=[0]*256
    a[1]=1 if pos==0 else 0
    a[2]=[0,2,3,0][pos]
    a[0]=[0,0,2,5][pos]
    a[3]=[0,0,0,2][pos]
    a[4]=[0,1,2,2][pos]
    records[(tno,8,12)]=a
    b=[100000*pos+j for j in range(256)];b[0]=123456+pos;b[255]=987654+pos
    records[(tno,9,12)]=b
records[(103,8,13)]=[0]*256
records[(104,8,13)]=[1]*256
records[(102,8,13)]=[0]*256
def record(tno,x,y,values,multiline=False):
    if multiline:
        return f'Tno={tno}\nX={x} Y={y}\nFBC BY BLK='+','.join(map(str,values[:100]))+'\n'+' '.join(map(str,values[100:]))+'\n'
    return f'Tno={tno} X={x} Y={y} FBC BY BLK='+' '.join(map(str,values))+'\n'
raw='HEADER: synthetic input, not production ASC\n'
for key in reversed(list(records)):
    raw+=record(*key,records[key],multiline=(key[0]==104))
raw+=record(103,8,12,[9]*256) # duplicate must not replace first
raw+=record(999,-9,-9,[4]*256) # unconfigured DUT must not be auto-discovered
EX=ROOT/'examples';EX.mkdir(exist_ok=True)
(EX/'valid_synthetic.asc').write_text(raw,encoding='ascii')
(EX/'second_file.asc').write_text(record(103,20,21,[7]*256),encoding='ascii')
def judge(dut,block,threshold):
    if not any((t,*dut) in records for t in seq):return 'No Data'
    cumulative=0
    for i,t in enumerate(seq):
        row=records.get((t,*dut))
        if row is None:return 'Missing Data'
        value=row[block];cumulative+=value
        if (value!=0 if threshold==0 else cumulative>=threshold):
            return 'Fail' if i==0 else volts[i-1]
    return 'Pass'
def expected(dut,threshold):
    result=[]
    for t,v in zip(seq,volts):
        result.append([t,v,*dut,*records.get((t,*dut),['Not Found']*256)])
    result.append([f'X{dut[0]}Y{dut[1]} Vmin','','','',*[judge(dut,b,threshold) for b in range(256)]])
    return result
all_expected={}
for threshold in (0,1,5,100):
    for dut in ((8,12),(9,12),(8,13),(99,99)):
        name=f'expected_{threshold}_{dut[0]}_{dut[1]}.tsv'
        rows=expected(dut,threshold)
        all_expected[name]=rows
        (EX/name).write_text('\n'.join('\t'.join(map(str,row)) for row in rows)+'\n',encoding='ascii')
(EX/'expected_results.json').write_text(json.dumps(all_expected,indent=2),encoding='utf-8')
def adapt(text):
    text=text.replace('Private Sub','Sub').replace('Private Function','Function')
    text=re.sub(r'ByRef (cache|present|volts|tnos)\(\)',r'ByRef \1',text)
    text=re.sub(r' As (String|Variant|Long|Double|Boolean|Object)\b','',text)
    text=re.sub(r'\b(\w+)\$',r'\1',text);text=re.sub(r'(\d)#',r'\1',text)
    text=text.replace('(0 To ','(').replace('(1 To ','(').replace(', 1 To ',', ')
    text=text.replace('LBound(present)','1')
    return re.sub(r'\bNext \w+','Next',text)
helpers=source[source.index('Private Sub ResetParser'):source.index('Private Function SortedDUTs')]
helpers+=source[source.index('Private Sub SortXY'):source.index('Private Sub WriteGroup')]
tests=r'''
Const MAX_EXACT=9007199254740991
Dim mTags,mTagHeads,mBlockTag,mSpaces,mDigits,mData,mWanted,mDUTs
Dim mTno,mX,mY,mPath,mActive,mRecordKey,mRecordDUT,mValues(),mCount,mLine,mRecordLine
Dim fso,folder,raw,buffer,chunk,p,q,line,tnos(4),volts(4),result,threshold,dut,xy,expectedFile,r,c,fields
Dim saved,bad,keys,a,i,value,cache(4),present(4),j,done
Set fso=CreateObject("Scripting.FileSystemObject"):folder=WScript.Arguments(0)
Set mTags=Nothing:Set mDigits=Nothing
Set mWanted=CreateObject("Scripting.Dictionary")
tnos(1)="103":tnos(2)="101":tnos(3)="104":tnos(4)="102"
For i=1 To 4:mWanted.Add tnos(i),True:Next
volts(1)=0.60:volts(2)=0.65:volts(3)=0.70:volts(4)=0.75
raw=fso.OpenTextFile(fso.BuildPath(folder,"valid_synthetic.asc")).ReadAll
For Each chunk In Array(1,7,262144)
 ResetParser "fixture":buffer=""
 For p=1 To Len(raw) Step chunk
  buffer=buffer & Mid(raw,p,chunk)
  Do
   q=InStr(buffer,vbLf):If q=0 Then Exit Do
   mLine=mLine+1:ParseLine Replace(Left(buffer,q-1),vbCr,"")
   buffer=Mid(buffer,q+1)
  Loop
 Next
 If Len(buffer)>0 Then ParseLine buffer
 FinishRecord
 If mData.Count<>11 Or mDUTs.Count<>3 Or mDUTs.Exists("-9|-9") Then Die "Record count / configured DUT discovery"
 a=mData("103|8|12"):If a(0)<>0 Or a(1)<>1 Then Die "First duplicate not retained"
 a=mData("103|9|12"):If a(0)<>123456 Or a(255)<>987654 Then Die "First or last FBC shifted"
 For Each threshold In Array(0,1,5,100)
  For Each dut In Array("8|12","9|12","8|13","99|99")
   result=BuildDUTGroup(mData,tnos,volts,dut,threshold)
   xy=Split(dut,"|")
   Set expectedFile=fso.OpenTextFile(fso.BuildPath(folder,"expected_" & threshold & "_" & xy(0) & "_" & xy(1) & ".tsv"))
   For r=1 To 5
    fields=Split(expectedFile.ReadLine,vbTab)
    For c=1 To 260
     If IsNumeric(fields(c-1)) And Len(fields(c-1))>0 Then
      If Not IsNumeric(result(r,c)) Then Die "Expected numeric cell at " & r & "," & c
      If CDbl(result(r,c))<>CDbl(fields(c-1)) Then Die "Numeric mapping/judge mismatch at " & r & "," & c
     Else
      If CStr(result(r,c))<>fields(c-1) Then Die "Text/missing mapping mismatch at " & r & "," & c
     End If
    Next
   Next
   expectedFile.Close
  Next
 Next
Next
WScript.Echo "PASS parser/output: chunks 1/7/262144; 256 raw indices; multiple DUT/TNO; nonnumeric TNO order; first duplicate; Legacy and thresholds 1/5/100; Missing Data and No Data; every output cell checked against independent Python oracle."
keys=Array("9|1","8|13","8|12","-2|5","9|-1")
SortXY keys,0,UBound(keys)
If Join(keys,",")<>"-2|5,8|12,8|13,9|-1,9|1" Then Die "Numeric XY sort"
ResetParser "second":ParseLine "Tno=103 X=20 Y=21 FBC BY BLK=" & RepeatValue("7",256):FinishRecord
If mData.Count<>1 Or mData.Exists("103|8|12") Then Die "Cross-file contamination"
For Each bad In Array(RepeatValue("0",255),RepeatValue("0",257),RepeatValue("0",255) & " BAD",RepeatValue("0",255) & " -1")
 ResetParser "bad"
 On Error Resume Next
 ParseLine "Tno=103 X=8 Y=12 FBC BY BLK=" & bad
 FinishRecord:saved=Err.Number:Err.Clear
 On Error GoTo 0
 If saved=0 Then Die "Malformed FBC accepted"
Next
ResetParser "bad context":ParseLine "Tno=103 X=8 Y=12"
On Error Resume Next
ParseLine "Tno=oops":saved=Err.Number:Err.Clear
On Error GoTo 0
If saved=0 Then Die "Malformed tag reused old TNO"
For Each bad In Array("-1","1.5","abc","")
 On Error Resume Next
 value=UnsignedValue(bad,MAX_EXACT):saved=Err.Number:Err.Clear
 On Error GoTo 0
 If saved=0 Then Die "Invalid threshold accepted"
Next
' A trigger before a later missing record remains determined; missing before trigger cannot be zero.
For i=1 To 4:present(i)=True:cache(i)=Split(RepeatValue("0",256)," "):Next
present(4)=False:a=cache(2):a(0)=1:cache(2)=a
If JudgeBlock(cache,present,volts,0,0)<>0.60 Then Die "Later missing row invalidated determined trigger"
If JudgeBlock(cache,present,volts,1,0)<>"Missing Data" Then Die "Missing row became zero/Pass"
WScript.Echo "PASS numeric DUT sorting, per-file reset, malformed 255/257/noninteger/negative FBC, stale tag rejection, threshold validation and missing-data edge cases."
Function RepeatValue(value,count)
 Dim values(),index:ReDim values(count-1)
 For index=0 To count-1:values(index)=value:Next
 RepeatValue=Join(values," ")
End Function
Sub Die(message)
 WScript.Echo "FAIL: " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
    path=Path(folder)/'test.vbs';path.write_text(adapt(tests+'\n'+helpers),encoding='ascii')
    result=subprocess.run(['cscript.exe','//nologo',str(path),str(EX)],capture_output=True,check=True)
    output=(result.stdout+result.stderr).decode('cp950',errors='replace');print(output.encode('ascii','backslashreplace').decode())
    if 'PASS numeric DUT sorting' not in output:
        match=re.search(r'\((\d+),\s*\d+\)',output)
        if match:
            number=int(match.group(1));lines=path.read_text().splitlines()
            for index in range(max(0,number-3),min(len(lines),number+2)):
                print(f'{index+1}: {lines[index]}')
        raise RuntimeError('VBA logic tests did not finish')
