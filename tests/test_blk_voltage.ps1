$ErrorActionPreference='Stop';$root='C:\Users\dora9\outputs\ASC_ThreeASC_Project';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
try{
$app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
$book=$app.Workbooks.Open("$root\Latest\ASC_BLK_Voltage_Analysis_Demo64.xlsm",0,$false)
$map=$book.Worksheets.Item('Wafer Maps');$idx=$book.Worksheets.Item('Import Index');$panelTop=[int]$idx.Cells.Item(2,50).Value2
$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_OpenIndexedMap",$book,$map,$map.Cells.Item(($panelTop+2),8))
$d=$book.Worksheets.Item('Chip Fail Details')
Check ([string]$d.Range('A6').Value2 -like 'BLK0*') 'BLK0 position'
Check ([string]$d.Range('A21').Value2 -like 'BLK15*') 'BLK15 position'
Check ([string]$d.Range('B6').Value2 -like 'BLK16*') 'BLK16 position'
Check ([string]$d.Range('P21').Value2 -like 'BLK255*') 'BLK255 position'
$metadata=$idx.Range('BC2:BR257').Value2
Check ($metadata.GetValue(1,6) -eq 'Fail') 'BLK0 state'
Check ([string]$metadata.GetValue(1,7) -eq '503') ('First TNO '+$metadata.GetValue(1,7))
Check ($metadata.GetValue(1,9) -eq 'Equal' -and $metadata.GetValue(1,12) -eq 1 -and $metadata.GetValue(1,13) -eq 4 -and $metadata.GetValue(1,14) -eq 5) 'Observed comparison or missing evidence wrong'
Check ([string]$d.Range('A6').Value2 -like '*Equal*') 'Grid tendency'
$res=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_ShowBLKVoltage",$book,$d,$d.Range('A6'));Check ($res -like 'BLK0 details ready') $res
Check ($d.ListObjects.Item('tblBLKObservedWR').ListRows.Count -eq 5) 'Observed Addresses lost'
Check ($d.ListObjects.Item('tblBLKVoltage').ListRows.Count -eq 34) 'Function sequence length'
$chart=$d.ChartObjects().Item('blkVoltageChart').Chart;Check ($chart.SeriesCollection().Count -eq 1) 'March excluded BLK plotted';Check ($chart.SeriesCollection().Item(1).Name -eq 'Scan') 'Wrong Pattern series'
$t=$d.ListObjects.Item('tblBLKObservedWR');for($i=1;$i -le $t.ListRows.Count;$i++){$sr=[int]$t.DataBodyRange.Cells.Item($i,12).Value2;Check ($t.DataBodyRange.Cells.Item($i,12).Hyperlinks.Item(1).SubAddress -eq "'Fail Address'!A$sr") 'Address source row wrong'}
# A second BLK updates existing charts/tables instead of appending stale copies.
$res=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_ShowBLKVoltage",$book,$d,$d.Range('B6'));Check ($res -like 'BLK16 details ready') $res;Check ($d.ChartObjects().Count -eq 1) 'Duplicate voltage charts'
$res=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_ShowBLKVoltage",$book,$d,$d.Range('A6'));Check ($res -like 'BLK0 details ready') $res
$d.Activate();$app.Goto($d.Range('A1'),$true);$book.Save()
$range=$d.Range('A1:P25');$range.CopyPicture(1,2);$pic=$d.ChartObjects().Add(0,0,$range.Width,$range.Height);$pic.Activate();$pic.Chart.Paste();$pic.Chart.Export("$root\wafer128\BLK_Voltage_Grid_Preview.png",'PNG')|Out-Null;$pic.Delete()
Write-Output 'PASS native Excel:column-major0/15/16/255;firstTNO503;Equal1/Unknown4;34Function rows;Scan only when March excluded;5 exact Address links;chart reused across BLKs.'
}finally{if($book){$book.Close($false)};if($app){$app.Quit()}}
