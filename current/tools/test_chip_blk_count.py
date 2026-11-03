from pathlib import Path
import re,subprocess,tempfile
r=Path(__file__).resolve().parents[1];s=(r/'src/BLKVminEngine.bas').read_text()
a=s.index('Private Function CountFailBlocks');helper=s[a:s.index('End Function',a)+len('End Function')]
helper=helper.replace('Private Function','Function')
helper=re.sub(r' As (Variant|Long)\b','',helper);helper=re.sub(r'\bNext \w+','Next',helper)
code='''
Dim values(255),b,result
For b=0 To 255:values(b)=0:Next
If CountFailBlocks(values)<>0 Then Die "All zero is not0"
values(0)=1:values(255)=999999:values(17)=2
If CountFailBlocks(values)<>3 Then Die "FBC magnitude wrongly summed / boundary BLK lost"
For b=0 To 255:values(b)=1:Next
If CountFailBlocks(values)<>256 Then Die "Full256 count"
WScript.Echo "PASS Chip BLK Fail Count: BLK0 and255, all zero=0, positive blocks counted once regardless of FBC magnitude, max256."
Sub Die(message)
 WScript.Echo "FAIL " & message:WScript.Quit 1
End Sub
'''
with tempfile.TemporaryDirectory() as folder:
 p=Path(folder)/'count.vbs';p.write_text(code+'\n'+helper,encoding='ascii')
 result=subprocess.run(['cscript.exe','//nologo',str(p)],capture_output=True)
 out=(result.stdout+result.stderr).decode('cp950',errors='replace');print(out.encode('ascii','backslashreplace').decode())
 assert 'PASS Chip BLK Fail Count:' in out,out
