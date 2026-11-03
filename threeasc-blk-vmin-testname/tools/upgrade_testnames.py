"""Build the Test Name variant from preserved v0.13 source snapshots."""
from pathlib import Path
root=Path(__file__).resolve().parents[1]
main=(root/'reference/IntegratedMain_v013.bas').read_text()
main=main.replace('BLKVminEngine.BLK_WritePhase(ThisWorkbook.Worksheets("WT" & phase & "_BLK_Distribution"))','BLKVminEngine.BLK_WritePhase(ThisWorkbook, phase)')
for row in range(2,7):main=main.replace(f'cfg.Range("D{row}")',f'cfg.Range("F{row}")')
main=main.replace('    cfg.Range("Q2").Value2 = blkWarnings','    blkMessage = BLKVminEngine.BLK_Finalize(ThisWorkbook)\n    If Len(blkMessage) > 0 Then blkWarnings = blkWarnings & vbCrLf & blkMessage\n    cfg.Range("Q2").Value2 = blkWarnings')
(root/'src/ImportThreeASC_BLKVmin.bas').write_text(main)
engine=(root/'reference/BLKVminEngine_v013.bas').read_text()
engine=engine.replace('Private mThreshold As Double', 'Private mTestNames() As String, mGroups As Object, mSheetNames As Object, mNextRows As Object\nPrivate mThreshold As Double')
engine=engine.replace('ReadConfiguration cfg, mSequence, mVoltage, mConfiguredDUTs, mThreshold','ReadConfiguration cfg, mSequence, mTestNames, mVoltage, mConfiguredDUTs, mThreshold\n    Set mGroups = BuildNameGroups(mTestNames)\n    SetupNameSheets cfg.Parent')
a=engine.index('Public Function BLK_WritePhase');b=engine.index('Public Sub BLK_Release',a)
engine=engine[:a]+(root/'tools/parts/testname_writer.vba').read_text()+'\n'+engine[b:]
engine=engine.replace('    mCarry = "": mActive = False\nEnd Sub','    mCarry = "": mActive = False\n    Set mGroups = Nothing: Set mSheetNames = Nothing: Set mNextRows = Nothing\nEnd Sub',1)
a=engine.index('Private Sub ReadConfiguration');b=engine.index('Private Sub ResetParser',a)
config=engine[a:b]
config=config.replace('ByRef tnos() As String, ByRef volts()', 'ByRef tnos() As String, ByRef names() As String, ByRef volts()')
config=config.replace('Config!A2:B','Config!A2:C').replace('"A2:B"','"A2:C"')
config=config.replace('TNO and voltage','TNO, Test Name and voltage')
config=config.replace('data(i, 2)','data(i, 3)')
config=config.replace('If IsError(data(i, 1)) Or IsError(data(i, 3)) Then','If IsError(data(i, 1)) Or IsError(data(i, 2)) Or IsError(data(i, 3)) Then')
config=config.replace('ReDim volts(1 To UBound(data, 1))','ReDim volts(1 To UBound(data, 1)): ReDim names(1 To UBound(data, 1))')
config=config.replace('            count = count + 1:', '            If Len(Trim$(CStr(data(i, 2)))) = 0 Then Err.Raise vbObjectError + 401, , "Test Name missing at Config row " & i + 1\n            count = count + 1: names(count) = Trim$(CStr(data(i, 2))):')
config=config.replace('ElseIf Len(Trim$(CStr(data(i, 3)))) > 0 Then', 'ElseIf Len(Trim$(CStr(data(i, 3)))) > 0 Or Len(Trim$(CStr(data(i, 2)))) > 0 Then')
config=config.replace('Voltage without TNO', 'Test Name/voltage without TNO')
config=config.replace('ReDim Preserve volts(1 To count)','ReDim Preserve volts(1 To count): ReDim Preserve names(1 To count)')
engine=engine[:a]+config+engine[b:]
engine+='\n'+(root/'tools/parts/testname_helpers.vba').read_text()
(root/'src/BLKVminEngine.bas').write_text(engine)
base=(root/'reference/BLKVminEngine_v013.bas').read_text()
assert engine[engine.index('Private Sub ResetParser'):engine.index('Private Sub WriteGroup')]==base[base.index('Private Sub ResetParser'):base.index('Private Sub WriteGroup')]
print('PASS: ABC Config, ordered Test Name groups, per-name sheets with three WT sections; FBC parser/Judge unchanged.')
