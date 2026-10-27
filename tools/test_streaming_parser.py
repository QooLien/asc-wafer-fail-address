"""Exercise actual BAS parser in VBScript. Does not test Excel UI or VBA compilation."""
import argparse
import pathlib
import re
import subprocess
import tempfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
source = (ROOT / 'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
helpers = source[source.index('Private Sub ParseASCBuffer'):]
helpers = re.sub(r'Private Sub FlushImportLog\b.*?End Sub\n', '', helpers, flags=re.S)
helpers = helpers.replace('Private Function', 'Function').replace('Private Sub', 'Sub')
helpers = re.sub(r' As (String|Variant|Long|Double|Boolean|Object)\b', '', helpers)
helpers = re.sub(r'\b(\w+)\$', r'\1', helpers)
helpers = re.sub(r'(\d)#', r'\1', helpers).replace('(0 To 7)', '(7)')
helpers = re.sub(r'\bNext \w+', 'Next', helpers)
helpers = helpers.replace('    On Error GoTo Invalid\n', '').replace('Invalid:\n    IntegerInRange = False\n', '')
cases = (ROOT / 'tools/address_cases.vbs').read_text(encoding='utf-8').strip()
assert len(cases.splitlines()) == 20
fixture = ('tno= 100 dbk=1 > [ 2, 100, 10, 14, 0, 25, 3, 4, ] [2,100,10,14,0,25,3,4,]\n'
           '[2,100,11,14,255,1023,15,15,] [2,100,10,14,256,0,0,0,]\n'
           'tno=101 dbk=1 > [2,101,10,14,0,25,3,4,] [2,100,12,14,0,0,0,0,]\n'
           'tno=999 dbk=1 > [2,999,10,14,0,0,0,0,]')
fixture_vbs = ' & vbLf & '.join('"' + line + '"' for line in fixture.splitlines())
tests = r'''
Const MAX_PENDING_RECORD = 16384
Dim mIntegerRegex
Set mIntegerRegex = Nothing
Dim nums, reason, tnos, addresses, flags, logs, records, runID, path
Dim sourceLine, headerTNO, selectedSections, skipped, accepted, buffer, entry, raw
Dim p, chunkSize, v, record, count, maxBuffer, file, chunk, startTime, total
Set tnos = CreateObject("Scripting.Dictionary")
tnos.Add "100", 7: tnos.Add "101", 8
runID = "test": path = "synthetic.asc"
__CASES__
raw = __FIXTURE__
For Each chunkSize In Array(1,7,65536)
 ResetState
 For p = 1 To Len(raw) Step chunkSize
  buffer = buffer & Mid(raw, p, chunkSize)
  Pump False
 Next
 Pump True
 If addresses.Count <> 2 Or accepted <> 4 Or skipped <> 2 Or selectedSections <> 2 Then Die "Counts with chunk size " & chunkSize
 If Not flags.Exists("10,14,0,25,3,4|100") Or Not flags.Exists("10,14,0,25,3,4|101") Then Die "TNO flags"
 If logs.Count <> 2 Then Die "Log count"
 entry = logs.Item(0)
 If entry(2) <> "WARN" Or entry(5) <> "100" Or entry(6) <> 2 Or entry(8) <> "BLK=256; allowed 0-255." Then Die "Range log context"
 entry = logs.Item(1)
 If entry(6) <> 3 Or entry(8) <> "Bracket TNO differs from section TNO." Then Die "TNO mismatch log"
Next
ResetState
buffer = "tno=100 > [2,100": Pump True
If skipped <> 1 Then Die "Missing closing bracket"
entry = logs.Item(0)
If entry(8) <> "Unterminated bracket at end of file." Then Die "EOF log"
ResetState
buffer = "tno=100 > [" & String(20000, " "): Pump False
If skipped <> 1 Or Len(buffer) <> 0 Then Die "Oversized pending buffer"
ResetState
buffer = "tno=100": Pump True
If selectedSections <> 1 Then Die "EOF header"
ResetState
buffer = "tno=100": Pump False
buffer = buffer & "0 > [2,1000,10,14,0,0,0,0,]": Pump True
If selectedSections <> 0 Or accepted <> 0 Then Die "Split non-selected TNO"
AddImportLog logs, runID, "ERROR", "Read ASC chunk", path, "100", 0, 53, "File not found.", ""
entry = logs.Item(0)
If entry(2) <> "ERROR" Or entry(7) <> 53 Then Die "Error log"
WScript.Echo "PASS: " & count & " address cases; chunks 1/7/65536; split headers/records; EOF/oversize guards; deduplication; log source lines and reasons."
If WScript.Arguments.Count > 0 Then
 ResetState
 Set file = CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.Arguments(0), 1, False, 0)
 startTime = Timer: total = 0: maxBuffer = 0
 Do Until file.AtEndOfStream
  chunk = file.Read(65536): total = total + Len(chunk)
  buffer = buffer & chunk: Pump False
  If Len(buffer) > maxBuffer Then maxBuffer = Len(buffer)
 Loop
 file.Close: Pump True
 If addresses.Count <> 2 Or accepted <> 4 Or skipped <> 2 Then Die "Large-file results"
 If maxBuffer > 16384 Then Die "Unbounded retained buffer"
 WScript.Echo "PASS large synthetic file: bytes=" & total & "; max retained buffer=" & maxBuffer & "; seconds=" & Round(Timer-startTime, 2)
End If
Sub ResetState
 Set addresses = CreateObject("Scripting.Dictionary")
 Set flags = CreateObject("Scripting.Dictionary")
 Set logs = CreateObject("Scripting.Dictionary")
 Set records = CreateObject("Scripting.Dictionary")
 buffer="": headerTNO="": sourceLine=1: selectedSections=0: skipped=0: accepted=0
End Sub
Sub Pump(eof)
 ParseASCBuffer buffer, eof, headerTNO, sourceLine, selectedSections, skipped, tnos, logs, runID, path, records
 For Each v In records.Keys
  record = records(v): addresses(record(0)) = True
  flags(record(0) & "|" & record(1)) = 1
 Next
 accepted = accepted + records.Count
End Sub
Sub Die(message)
 WScript.Echo "FAIL: " & message: WScript.Quit 1
End Sub
Sub Check(bracket, expected)
 Dim actual
 actual = TryParseAddress(bracket, nums, reason): count = count + 1
 If actual <> expected Then Die bracket & " " & reason
End Sub
'''.replace('__CASES__', cases).replace('__FIXTURE__', fixture_vbs)

cli = argparse.ArgumentParser()
cli.add_argument('--large-kb', type=int, default=0)
args = cli.parse_args()
with tempfile.TemporaryDirectory() as directory:
    directory = pathlib.Path(directory)
    script = directory / 'stream_test.vbs'
    script.write_text(tests + '\n' + helpers, encoding='ascii')
    command = ['cscript.exe', '//nologo', str(script)]
    if args.large_kb:
        sample = directory / 'large-synthetic.asc'
        padding = ('tno=999 >\n' + 'padding ' * 8190 + '\n').encode('ascii')
        with sample.open('wb') as output:
            remaining = args.large_kb * 1024
            while remaining:
                block = padding[:min(len(padding), remaining)]
                output.write(block)
                remaining -= len(block)
            output.write(b'\n' + fixture.encode('ascii'))
        command.append(str(sample))
    subprocess.run(command, check=True)
