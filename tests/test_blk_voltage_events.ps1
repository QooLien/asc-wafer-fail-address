$ErrorActionPreference='Stop';$root='C:\Users\dora9\outputs\ASC_ThreeASC_Project';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
try{
$app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
$book=$app.Workbooks.Open("$root\wafer128\BLK_Voltage_Event_Harness.xlsm",0,$false)
$idx=$book.Worksheets.Item('Import Index');$top=[int]$idx.Cells.Item(2,50).Value2
$res=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureEvent",'Wafer Maps',('H'+($top+2)));Check ($res -eq 'Cancelled|Chip Fail Details') ('Map event '+$res)
$res=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureEvent",'Chip Fail Details','A6');Check ($res -eq 'Cancelled|Chip Fail Details') ('BLK event '+$res)
$d=$book.Worksheets.Item('Chip Fail Details');$table=$d.ListObjects.Item('tblBLKObservedWR');$cell=$table.DataBodyRange.Cells.Item(1,12);$expected=[int]$cell.Value2
$res=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureSelection",'Chip Fail Details',$cell.Address($false,$false));Check ($res -eq "Fail Address!A$expected") ('Source event '+$res)
# Missing Function voltage is not zero and cannot be reported as a valid fail voltage.
$cfg=$book.Worksheets.Item('Config');$cfgData=$cfg.Range('D2:F901').Value2;$cfgRow=0
for($i=1;$i -le 900;$i++){if([string]$cfgData.GetValue($i,1) -eq '503'){$cfgRow=$i+1;break}}
$originalV=$cfg.Cells.Item($cfgRow,6).Value2;$cfg.Cells.Item($cfgRow,6).ClearContents()
$res=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureEvent",'Wafer Maps',('H'+($top+2)));Check ($res -eq 'Cancelled|Chip Fail Details') $res
Check ($null -eq $idx.Range('BJ2').Value2 -and $idx.Range('BR2').Value2 -eq 'Function voltage missing') 'Missing Function voltage handled incorrectly'
$cfg.Cells.Item($cfgRow,6).Value2=$originalV
# Remove one earlier FBC record in this disposable fixture: first failure becomes uncertain; raw blank remains blank.
$fbc=$book.Worksheets.Item('FBC by BLK');$last=$fbc.Cells.Item(1048576,2).End(-4162).Row;$data=$fbc.Range("A2:D$last").Value2;$sourceRow=0
for($i=1;$i -le $data.GetLength(0);$i++){if($data.GetValue($i,1) -eq 'WT1' -and [string]$data.GetValue($i,2) -eq '503' -and $data.GetValue($i,3) -eq 10 -and $data.GetValue($i,4) -eq 7){$sourceRow=$i+1;break}}
Check ($sourceRow -gt 0) 'Fixture FBC record not found';$fbc.Cells.Item($sourceRow,2).Value2=999
$res=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureEvent",'Wafer Maps',('H'+($top+2)));Check ($res -eq 'Cancelled|Chip Fail Details') $res
Check ($null -eq $idx.Range('BJ2').Value2 -and $idx.Range('BR2').Value2 -like 'Earlier TNO missing*') 'Missing earlier TNO falsely yielded first voltage'
$res=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureEvent",'Chip Fail Details','A6');Check ($res -eq 'Cancelled|Chip Fail Details') $res
$d=$book.Worksheets.Item('Chip Fail Details');$t=$d.ListObjects.Item('tblBLKVoltage');$rows=$t.DataBodyRange.Value2;$found=$false
for($i=1;$i -le $rows.GetLength(0);$i++){if($rows.GetValue($i,1) -eq 'Scan' -and [string]$rows.GetValue($i,2) -eq '503'){Check ($null -eq $rows.GetValue($i,5)) 'Missing FBC became zero';$found=$true}}
Check $found 'Missing TNO row omitted';Check ($d.ChartObjects().Count -eq 1) 'Chart duplicates'
Write-Output 'PASS actual Workbook events:Map double-click→grid;BLK double-click→chart;Source Row selection→Fail Address;missing Function voltage and earlier FBC do not become zero or certain first-fail voltage.'
}finally{if($book){$book.Close($false)};if($app){$app.Quit()}}
