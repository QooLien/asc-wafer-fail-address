$ErrorActionPreference='Stop';$root='C:\Users\dora9\outputs\ASC_ThreeASC_Project';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
try{
$app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
$book=$app.Workbooks.Open("$root\Latest\ASC_BLK_Voltage_Analysis_Demo64.xlsm",0,$false)
$map=$book.Worksheets.Item('Wafer Maps');$idx=$book.Worksheets.Item('Import Index');$panels=$idx.Range('AT2:BA19').Value2;$lookup=@{}
for($i=1;$i -le 6;$i++){$lookup[[string]$panels.GetValue($i,3)+'|'+$panels.GetValue($i,4)]=[int]$panels.GetValue($i,5)}
$rows=Import-Csv "$root\wafer128\Expected_BLK_Voltage_Grid.csv";$groups=$rows | Group-Object WT,Pattern,X,Y;$verified=0
foreach($g in $groups){
$r=$g.Group[0];$top=$lookup[$r.WT+'|'+$r.Pattern];$cell=$map.Cells.Item(($top+2+2*([int]$r.Y-7)),(2+2*([int]$r.X-7)))
$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_OpenIndexedMap",$book,$map,$cell)
$d=$book.Worksheets.Item('Chip Fail Details');$m=$idx.Range('BC2:BR257').Value2
for($c=1;$c -le 16;$c++){Check ($d.Columns.Item($c).ColumnWidth -eq 12) ('Grid column width '+$c+'='+$d.Columns.Item($c).ColumnWidth)}
foreach($e in $g.Group){$j=[int]$e.BLK+1;Check ([string]$m.GetValue($j,6) -eq $e.State) ('State '+$g.Name+'/'+$e.BLK);Check ([string]$m.GetValue($j,7) -eq $e.FirstTNO) ('FirstTNO '+$g.Name+'/'+$e.BLK)
if($e.FirstVoltage -eq ''){Check ($null -eq $m.GetValue($j,8)) 'Missing voltage became zero'}else{Check ([math]::Abs([double]$m.GetValue($j,8)-[double]$e.FirstVoltage) -lt 0.0000001) 'Voltage differs'}
Check ([string]$m.GetValue($j,9) -eq $e.Tendency) 'Tendency differs';$ns=@('W','R','Equal','Unknown','Captured');for($k=0;$k -lt 5;$k++){Check ([int]$m.GetValue($j,(10+$k)) -eq [int]$e.($ns[$k])) ('W/R count '+$g.Name+'/'+$e.BLK+'/'+$ns[$k])};$verified++}
Write-Output ('PASS '+$g.Name+' all256 BLK states/first voltage/observed W-R counters.')
}
# Isolated test-only edits: both-side observed voltage, mixed classes and missing Monitor voltage.
$src=$book.Worksheets.Item('Monitor Fail Analysis');$data=$src.UsedRange.Value2;$selectedRows=@()
for($i=5;$i -le $data.GetLength(0);$i++){if($data.GetValue($i,1) -eq 'WT1' -and $data.GetValue($i,2) -eq 'Scan' -and $data.GetValue($i,4) -eq 10 -and $data.GetValue($i,5) -eq 7 -and $data.GetValue($i,6) -eq 0){$selectedRows+=$i}}
Check ($selectedRows.Count -eq 5) 'Fixture Address count';$new=@(@('350','371'),@('351','370'),@('351','371'),@('',''),@('360',''))
for($i=0;$i -lt 5;$i++){$src.Cells.Item($selectedRows[$i],11).Value2=$new[$i][0];$src.Cells.Item($selectedRows[$i],12).Value2=$new[$i][1]}
$top=$lookup['WT1|Scan'];$cell=$map.Cells.Item(($top+2),8);$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_OpenIndexedMap",$book,$map,$cell)
$m=$idx.Range('BC2:BR2').Value2;Check ($m.GetValue(1,9) -eq 'Mixed' -and $m.GetValue(1,10) -eq 1 -and $m.GetValue(1,11) -eq 1 -and $m.GetValue(1,12) -eq 1 -and $m.GetValue(1,13) -eq 2) 'Mixed observed classification wrong'
$cfg=$book.Worksheets.Item('Config');$cfgData=$cfg.Range('A2:C901').Value2;$monitorRow=0
for($i=1;$i -le 900;$i++){if([string]$cfgData.GetValue($i,1) -eq '371'){$monitorRow=$i+1;break}}
$cfg.Cells.Item($monitorRow,3).ClearContents();$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_OpenIndexedMap",$book,$map,$cell);$m=$idx.Range('BC2:BR2').Value2
Check ($m.GetValue(1,9) -eq 'Read worse' -and $m.GetValue(1,11) -eq 1 -and $m.GetValue(1,13) -eq 4) 'Missing Monitor voltage misclassified'
Write-Output "PASS $verified independent BLK comparisons; test-only W1/R1/Equal1/Unknown2 Mixed;missing Monitor voltage -> Unknown;fixture edits not saved."
}finally{if($book){$book.Close($false)};if($app){$app.Quit()}}
