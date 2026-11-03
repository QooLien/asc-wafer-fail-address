"""Apply dual item configuration after upgrade_testnames.py."""
from pathlib import Path
root=Path(__file__).resolve().parents[1]
p=root/'src/ImportThreeASC_BLKVmin.bas';main=p.read_text()
a=main.index('    settings = cfg.Range("A2:A901")');b=main.index('    stage = "Select three ASC files"',a)
main=main[:a]+'''    stage = "Validate Monitor / Function settings"
    BLKVminEngine.BLK_Setup cfg
    Set tnos = BLKVminEngine.BLK_AddressTNOs()
'''+main[b:]
main=main.replace('cfg.Range("G2:L4")','cfg.Range("J2:O4")').replace('cfg.Range("Q2")','cfg.Range("T2")')
for row in range(2,7):main=main.replace(f'cfg.Range("F{row}")',f'cfg.Range("I{row}")')
p.write_text(main)
p=root/'src/BLKVminEngine.bas';s=p.read_text()
s=s.replace('Private mThreshold As Double','Private mMonitorTNOs As Object\nPrivate mThreshold As Double')
a=s.index('    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',s.index('Private Sub ReadConfiguration'));b=s.index('    value = ws.Range("N2")',a)
s=s[:a]+'''    Set mMonitorTNOs = CreateObject("Scripting.Dictionary")
    ReDim tnos(1 To 900): ReDim names(1 To 900): ReDim volts(1 To 900)
    ReadItemColumns ws.Range("A2:C901").Value2, True, seen, tnos, names, volts, count
    ReadItemColumns ws.Range("D2:F901").Value2, False, seen, tnos, names, volts, count
    If count = 0 Then Err.Raise vbObjectError + 301, , "Enter Monitor or Function TNOs in Config!A:F."
    ReDim Preserve tnos(1 To count): ReDim Preserve names(1 To count): ReDim Preserve volts(1 To count)
'''+s[b:]
s=s.replace('ws.Range("N2")','ws.Range("Q2")').replace('ws.Rows.Count, 14','ws.Rows.Count, 17').replace('ws.Rows.Count, 15','ws.Rows.Count, 18').replace('ws.Range("N6:O"','ws.Range("Q6:R"')
s=s.replace('group = BuildDUTGroup(mData, tnos, volts, CStr(dut), mThreshold)','group = BuildItemDUTGroup(mData, tnos, volts, CStr(dut), mThreshold, mMonitorTNOs)')
s=s.replace('groupSize = UBound(tnos) + 1','groupSize = UBound(tnos) + IIf(HasMonitor(tnos, mMonitorTNOs), 1, 0)')
s=s.replace('    For i = 1 To duts.Count\n        block = dataRow', '    If HasMonitor(tnos, mMonitorTNOs) Then\n    For i = 1 To duts.Count\n        block = dataRow')
s=s.replace('    mNextRows(name) = endRow + 1','    End If\n    mNextRows(name) = endRow + 1',1)
# Comparison view: Function-only groups have no Vmin rows, while mixed names judge Monitor only.
s=s.replace('Dim color As Long, note As String','Dim color As Long, note As String, addVmin As Boolean')
s=s.replace('    n = UBound(tnos)\n    last = ws.UsedRange', '    n = UBound(tnos): addVmin = HasMonitor(tnos, mMonitorTNOs)\n    last = ws.UsedRange')
s=s.replace('CDbl(ordered.Count) * (n + 1) * 3','CDbl(ordered.Count) * (n + IIf(addVmin, 1, 0)) * 3')
s=s.replace('CStr(dut), funcName, tnos, volts)','CStr(dut), funcName, tnos, volts, addVmin)')
s=s.replace('        With ws.Cells(dest + rr - 3', '        If addVmin Then\n        With ws.Cells(dest + rr - 3')
s=s.replace('        dest = dest + rr','        End If\n        dest = dest + rr')
s=s.replace('ByRef volts() As Double) As Variant\n    Dim out() As Variant, xy As Variant, n As Long, rr As Long', 'ByRef volts() As Double, Optional ByVal addVmin As Boolean = True) As Variant\n    Dim out() As Variant, xy As Variant, n As Long, rr As Long')
s=s.replace('ReDim out(1 To (n + 1) * 3','ReDim out(1 To (n + IIf(addVmin, 1, 0)) * 3').replace('For i = 1 To n + 1\n            For phase', 'For i = 1 To n + IIf(addVmin, 1, 0)\n            For phase')
s+='\n'+(root/'tools/parts/item_groups.vba').read_text()
p.write_text(s)
print('PASS source generation: Monitor A:C, Function D:F; Monitor-only Vmin; union address index preserved.')
