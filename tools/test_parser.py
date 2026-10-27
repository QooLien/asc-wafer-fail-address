"""Run actual BAS parser helpers in Windows VBScript after syntax-only adaptation.

This tests the shared VBScript.RegExp engine and parsing logic, not Excel/VBA compilation.
"""
import pathlib
import re
import subprocess
import tempfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
source = (ROOT / 'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
helpers = source[source.index('Private Function TryParseAddress'):]
helpers = re.sub(r'Private Function', 'Function', helpers)
helpers = re.sub(r' As (String|Variant|Long|Double|Boolean|Object)\b', '', helpers)
helpers = re.sub(r'\b(\w+)\$', r'\1', helpers)
helpers = re.sub(r'(\d)#', r'\1', helpers)
helpers = helpers.replace('(0 To 7)', '(7)')
helpers = re.sub(r'\bNext \w+', 'Next', helpers)
# VBScript only supports On Error Resume Next; test inputs stay within Double range.
helpers = helpers.replace('    On Error GoTo Invalid\n', '')
helpers = helpers.replace('Invalid:\n    IntegerInRange = False\n', '')
loop = source[source.index('    For Each token In tokens'):source.index('    Next token') + len('    Next token')]
loop = re.sub(r'\b(\w+)\$', r'\1', loop)
loop = re.sub(r'\bNext \w+', 'Next', loop)
loop = loop.replace('    For Each token In tokens', '    For Each token In tokens')

tests = '''
Dim nums, reason, good, bad, tnos, addresses, raw, re, tokens, token
Dim s, headerTNO, selectedSections, valid, tno, key, flags, accepted, skipped, firstReject, count
Set tnos = CreateObject("Scripting.Dictionary")
Set addresses = CreateObject("Scripting.Dictionary")
tnos.Add "100", 7
tnos.Add "101", 8
Call Check("[ 2, 100, 10, 14, 0, 25, 3, 4, ]", True)
Call Check("[2,100,10,14,255,1023,15,15]", True)
Call Check("[2,100,10,14,0,0,0,0,]", True)
Call Check("[2," & vbTab & "100,10," & vbCrLf & "14,0,0,0,0," & vbCrLf & "]", True)
Call Check("[2," & ChrW(160) & "100,10,14,0,0,0,0," & ChrW(12288) & "]", True)
Call Check("[2,100,10,14,256,0,0,0,]", False)
Call Check("[2,100,10,14,0,1024,0,0,]", False)
Call Check("[2,100,10,14,0,0,16,0,]", False)
Call Check("[2,100,10,14,0,0,0,16,]", False)
Call Check("[2,100,10,14,-1,0,0,0,]", False)
Call Check("[2,100,10,14,0,0,0,-1,]", False)
Call Check("[2,99,10,14,0,0,0,0,]", False)
Call Check("[2,1000,10,14,0,0,0,0,]", False)
Call Check("[2,100,10,14,0,0,,0,]", False)
Call Check("[2,100,10,14,0,0,1.5,0,]", False)
Call Check("[2,100,10,14,0,0,1 2,0,]", False)
Call Check("[2,100,10,14,0,0,0,]", False)
Call Check("[2,100,10,14,0,0,0,0,9,]", False)
Call Check("[2,100,2147483648,14,0,0,0,0,]", False)
raw = "tno= 100 dbk=1 > [ 2, 100, 10, 14, 0, 25, 3, 4, ] [2,100,10,14,0,25,3,4,]" & vbCrLf & _
      "[2,100,11,14,255,1023,15,15,] [2,100,10,14,256,0,0,0,]" & vbCrLf & _
      "tno=101 dbk=1 > [2,101,10,14,0,25,3,4,] [2,100,12,14,0,0,0,0,]" & vbCrLf & _
      "tno=999 dbk=1 > [2,999,10,14,0,0,0,0,]"
Set re = CreateObject("VBScript.RegExp")
re.Global = True
re.IgnoreCase = True
re.Pattern = "\\btno\\s*=\\s*[0-9]+\\b|\\[[^\\]]*\\]"
Set tokens = re.Execute(raw)
'''
assertions = '''
If addresses.Count <> 2 Or accepted <> 4 Or skipped <> 2 Or selectedSections <> 2 Then
 WScript.Echo "FAIL integration: " & addresses.Count & "," & accepted & "," & skipped & "," & selectedSections
 WScript.Quit 1
End If
Set flags = addresses.Item("10,14,0,25,3,4")
If flags.Count <> 2 Or flags("100") <> 1 Or flags("101") <> 1 Then WScript.Quit 2
WScript.Echo "PASS: " & count & " address cases; multi-TNO sections, deduplication, and address flags."
Sub Check(bracket, expected)
 Dim actual
 actual = TryParseAddress(bracket, nums, reason)
 count = count + 1
 If actual <> expected Then
  WScript.Echo "FAIL " & bracket & " " & reason
  WScript.Quit 3
 End If
End Sub
'''
with tempfile.TemporaryDirectory() as tmp:
    path = pathlib.Path(tmp) / 'parser_test.vbs'
    path.write_text(tests + '\n' + loop + '\n' + assertions + '\n' + helpers, encoding='ascii')
    result = subprocess.run(['cscript.exe', '//nologo', str(path)])
    if result.returncode:
        for no, line in enumerate(path.read_text().splitlines(), 1):
            if 57 <= no <= 68:
                print(f'{no}: {line}')
        raise SystemExit(result.returncode)
