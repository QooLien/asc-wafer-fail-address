from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/ImportThreeASC_BLKVmin.bas').read_text()
# Legacy three-stage adapter; native_import_all tests actual1/2/3-file execution.
s=s.replace("For phase = 1 To fileCount - 1","For phase = 1 To 2").replace("For phase = 1 To fileCount","For phase = 1 To 3").replace("For other = phase + 1 To fileCount","For other = phase + 1 To 3")
a=s.index('Private Sub AddAddressLabels');helper=s[a:s.index('End Sub',a)+len('End Sub')]
a=s.index('Private Function AddressVoltageLabel');helper+='\n'+s[a:s.index('End Function',a)+len('End Function')]
helper=helper.replace('Private Sub','Sub').replace('Private Function','Function')
helper=re.sub(r' As (Variant|String|Long|Double|Object)\b','',helper)
helper=re.sub(r'\b(\w+)\$',r'\1',helper);helper=re.sub(r'(\d)#',r'\1',helper)
helper=re.sub(r'\bNext \w+','Next',helper)
a=s.index('    For phase = 1 To 3\n        Set localTNOs = PhaseTNOs(tnos, phase)',s.index('ReDim headers'))
header=s[a:s.index('    Next phase',a)+len('    Next phase')]
header=re.sub(r'\bNext \w+','Next',header)
phase=s[s.index('Private Function PhaseTNOs'):s.index('Private Function StepNumber')]
phase=phase.replace('Private Function','Function');phase=re.sub(r' As (Variant|Long|Object)\b','',phase);phase=re.sub(r'\bNext \w+','Next',phase)
code='''
Dim labels,data(4,3),more(1,3),i,v,tnos,localTNOs,addressLabels,phase,headers(1,18),idx
Set labels=CreateObject("Scripting.Dictionary")
data(1,1)=350:data(1,2)="W-Low pBL":data(1,3)=1.2
data(2,1)=351:data(2,2)="W-Low pBL":data(2,3)=1.035
data(3,1)=355:data(3,2)="W-Low pBL":data(3,3)=0.765
data(4,1)=900
AddAddressLabels data,labels
more(1,1)=370:more(1,2)="R-Low pBL":more(1,3)=0.99
AddAddressLabels more,labels
If labels("350")<>"W-Low pBL_1.20v" Then Die "Two decimals"
If labels("351")<>"W-Low pBL_1.035v" Or labels("355")<>"W-Low pBL_0.765v" Then Die "Three decimals lost"
If labels("370")<>"R-Low pBL_0.99v" Or labels("900")<>"900" Then Die "Second input group / missing metadata fallback"
If AddressVoltageLabel(1)<>"1.00v" Or AddressVoltageLabel(0.6)<>"0.60v" Or AddressVoltageLabel(0.68)<>"0.68v" Then Die "Trailing zero"
If AddressVoltageLabel(1.2345)<>"1.2345v" Then Die "Additional meaningful precision lost"
Set tnos=CreateObject("Scripting.Dictionary")
tnos.Add "355",0:tnos.Add "350",1:tnos.Add "351",2:tnos.Add "370",3
Set addressLabels=labels
'''+header+'''
For phase=1 To 3
 For Each v In tnos
  idx=7+(phase-1)*tnos.Count+tnos(v)
  If headers(1,idx)<>"WT" & phase & vbLf & labels(v) Then Die "Header shifted TNO index"
 Next
Next
WScript.Echo "PASS address labels: two/three decimal precision, both Config groups, fallback TNO, actual header loop and shuffled TNO-to-column mapping for all3 WTs."
Function IsError(value)
 IsError=(VarType(value)=10)
End Function
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'labels.vbs';p.write_text(code+'\n'+helper+'\n'+phase,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS address labels:' in out,out
