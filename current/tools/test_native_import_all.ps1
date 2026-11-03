param([Parameter(Mandatory=$true)][string]$Harness,[string]$Project='C:/Users/dora9/outputs/ASC_ThreeASC_Project')
$ErrorActionPreference='Stop';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 foreach($n in 1..3){
  $book=$app.Workbooks.Open($Harness,0,$false)
  $result=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureImportAll",$n)
  Check ($result -like 'Import and analysis complete.*') $result
  $address=$book.Worksheets.Item('Fail Address');$fbc=$book.Worksheets.Item('FBC by BLK');$locations=$book.Worksheets.Item('Chip Locations')
  $expectedAddress=@(0,1000,1002,1005)[$n];$expectedFbc=@(0,177,295,413)[$n]
  Check ($address.Cells.Item(1048576,1).End(-4162).Row -eq ($expectedAddress+1)) "Address count wrong for$n files"
  Check ($fbc.Cells.Item(1048576,2).End(-4162).Row -eq ($expectedFbc+1)) "FBC count wrong for$n files"
  Check ($book.Worksheets.Item('BLK_FBC_Comparison').UsedRange.Columns.Count -eq (3+59*$n)) 'Comparison contains unselected WT columns'
  Check ($address.Cells.Item(1,16384).End(-4159).Column -eq (6+59*$n)) 'Address contains unselected WT columns'
  $matrix=$address.UsedRange.Value2;$index=$book.Worksheets.Item('Import Index').Range("A2:D$($n*59+1)").Value2
  for($r=2;$r -le $matrix.GetLength(0);$r++){
   $blk=[int]$matrix.GetValue($r,3);$wl=[int]$matrix.GetValue($r,4)
   for($j=1;$j -le $index.GetLength(0);$j++){
    $wt=[string]$index.GetValue($j,1);$tno=[int]$index.GetValue($j,2);$column=[int]$index.GetValue($j,3)
    $expected=0
    if($wt -eq 'WT1' -and $tno -eq 530 -and (($blk -eq 0 -and $wl -lt 996) -or ($blk -eq 1 -and $wl -lt 2))){$expected=1}
    if($wt -eq 'WT1' -and $tno -eq 717 -and $blk -eq 6 -and $wl -lt 2){$expected=1}
    if($wt -eq 'WT2' -and $tno -eq 530 -and $blk -eq 4 -and $wl -lt 2){$expected=1}
    if($wt -eq 'WT3' -and (($tno -eq 530 -and $blk -eq 5 -and $wl -eq 0) -or ($tno -eq 717 -and $blk -eq 8 -and $wl -lt 2))){$expected=1}
    Check ([double]$matrix.GetValue($r,$column) -eq $expected) "Address flag wrong at$n files,row$r,$wt,TNO$tno"
   }
  }
  $report=$book.Worksheets.Item('Missing BLK Address');$coverage=@(0,4,6,8)[$n]
  Check ($report.Cells.Item(1048576,1).End(-4162).Row -eq ($coverage+4)) 'Audit did not run automatically'
  $summary=$book.Worksheets.Item('BLK Fail Count');Check ($summary.Cells.Item(1048576,1).End(-4162).Row -eq (8*$n+1)) 'Map contains unselected WTs'
  foreach($name in @('Scan Fail Map','March Fail Map')){Check ($book.Worksheets.Item($name).UsedRange.Columns.Count -eq (4*$n)) 'Map has leftover panels'}
  for($phase=1;$phase -le 3;$phase++){
   $status=$locations.Cells.Item(($phase+1),5).Value2
   if($phase -le $n){Check ($status -eq 'OK') 'Provided WT failed'}else{
    Check ([string]::IsNullOrEmpty($status)) 'Absent WT should not be shown'
    for($r=2;$r -le (8*$n+1);$r++){Check ($summary.Cells.Item($r,5).Value2 -ne "WT$phase") 'Absent WT found in summary'}
   }
  }
  $cfg=$book.Worksheets.Item('Config')
  foreach($i in @(5,6,8,9,10,11,12)){Check ($cfg.Range('H1:N901').Borders.Item($i).LineStyle -eq -4142) ('Config H:N border'+$i+'='+$cfg.Range('H1:N901').Borders.Item($i).LineStyle)}
  Check ($book.Worksheets.Item($book.Worksheets.Count).Name -eq 'Log' -and $book.Worksheets.Item(($book.Worksheets.Count-1)).Name -eq 'Chip Locations') 'Last two sheet order wrong'
  Check ($cfg.Buttons().Item(1).Caption -eq '匯入並分析') 'Main button caption wrong'
  $book.SaveCopyAs((Join-Path $Project "wafer128/Native_Import_All_${n}WT.xlsm"))
  Write-Output "PASS one-button$n files: Address$expectedAddress, FBC$expectedFbc, automatic comparison/Map/coverage$coverage rows, ONLY imported WTs visible, Config H:N borders cleared, Chip Locations then Log last."
  $book.Close($false);$book=$null
 }
 $book=$app.Workbooks.Open((Join-Path $Project 'wafer128/Native_Import_All_3WT.xlsm'),0,$false)
 $result=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureImportAll",1)
 Check ($result -like 'Import and analysis complete.*') $result
 Check ($book.Worksheets.Item('Fail Address').Cells.Item(1,16384).End(-4159).Column -eq 65) 'Rerun3-to1 retained Address WT2/3'
 Check ($book.Worksheets.Item('BLK_FBC_Comparison').UsedRange.Columns.Count -eq 62) 'Rerun3-to1 retained comparison WT2/3'
 Check ($book.Worksheets.Item('Scan Fail Map').UsedRange.Columns.Count -eq 4) 'Rerun3-to1 retained Map panels'
 Write-Output 'PASS rerun3 WT to1 WT: stale Address/comparison/Map WT2/3 removed.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
