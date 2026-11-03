from pathlib import Path
import re
r=Path(__file__).resolve().parents[1];s=(r/'src/ImportThreeASC_BLKVmin.bas').read_text()
assert re.search(r'Public Sub ImportASCFailAddresses\(\)\s+RunThreeASC False\s+End Sub',s)
assert re.search(r'Public Sub ImportASCFBC\(\)\s+RunThreeASC True\s+End Sub',s)
body=s[s.index('Private Sub RunThreeASC'):s.index("' Only orchestration helpers")]
stack=[];address_lines=0;fbc_lines=0
for line in body.splitlines():
 text=line.strip()
 m=re.match(r'If (.+) Then$',text)
 if m:
  cond=m[1].lower();stack.append(True if cond=='fbconly' else False if cond=='not fbconly' else None)
  continue
 if text=='Else':
  if stack[-1] is not None:stack[-1]=not stack[-1]
  continue
 if text=='End If':stack.pop();continue
 if 'dst.' in text or 'Set dst =' in text or 'ParseASCBuffer ' in text or 'ReDim bits(' in text:
  address_lines+=1
  assert False in stack or text.startswith('If Not fbcOnly Then '),text
 if any(token in text for token in ['BLKVminEngine.BLK_Feed ','BLKVminEngine.BLK_WritePhase(', 'BLKVminEngine.BLK_Finalize(', 'BLKVminEngine.BLK_BeginFile ']):
  fbc_lines+=1
  assert True in stack or text.startswith('If fbcOnly Then '),text
assert not stack
assert address_lines>0 and fbc_lines>0
engine=(r/'src/BLKVminEngine.bas').read_text()
assert 'Worksheets("Fail Address")' not in engine and 'ParseASCBuffer' not in engine
assert 'Private busy As Boolean' in s and 'If busy Then' in body
map_entry=s[s.index('Public Sub BuildFailMaps'):s.index('Private Sub RunThreeASC')]
assert 'If busy Then' in map_entry and 'BLK_BuildFailMaps(ThisWorkbook)' in map_entry
print('PASS separate actions: independent public entry points; every address parse/write/allocation is Address-only; every FBC feed/write is FBC-only; shared busy guard; FBC engine has no Address worksheet access.')
