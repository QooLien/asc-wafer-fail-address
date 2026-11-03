param([string]$Harness='C:/Users/dora9/outputs/ASC_ThreeASC_Project/wafer128/Home_Fixture_Harness.xlsm',[string]$Project='C:/Users/dora9/outputs/ASC_ThreeASC_Project')
$ErrorActionPreference='Stop'
function Check($ok,$message){if(-not $ok){throw $message}}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Harness,0,$false)
 $msg=$app.Run("'"+$book.Name+"'!ImportFailAddresses.HomeFixtureRun",$false);Check ($msg -like 'Import complete.*') $msg
 $address=$book.Worksheets.Item('Fail Address');Check ($address.Cells.Item(1048576,1).End(-4162).Row -eq 1006) 'Expected1005 unique Address rows'
 Check ($address.Cells.Item(1,16384).End(-4159).Column -eq 183) 'Expected6+59x3 columns'
 $msg=$app.Run("'"+$book.Name+"'!ImportFailAddresses.HomeFixtureRun",$true);Check ($msg -like 'FBC import complete.*') $msg
 $fbc=$book.Worksheets.Item('FBC by BLK');Check ($fbc.Cells.Item(1048576,2).End(-4162).Row -eq 414) 'Expected413 FBC rows'
 $locations=$book.Worksheets.Item('Chip Locations');Check ($locations.Range('A2:B5').Rows.Count -eq 4) 'Expected4 chips'
 for($r=2;$r -le 4;$r++){Check ($locations.Cells.Item($r,5).Value2 -eq 'OK') 'FBC parser error'}
 $summary=$book.Worksheets.Item('BLK Fail Count');$observed=@{}
 for($r=2;$r -le 25;$r++){
  $key=[string]$summary.Cells.Item($r,1).Value2+'|'+$summary.Cells.Item($r,5).Value2+'|'+$summary.Cells.Item($r,6).Value2+'|'+$summary.Cells.Item($r,7).Value2
  $observed[$key]=@($summary.Cells.Item($r,8).Value2,$summary.Cells.Item($r,9).Value2)
 }
 Check ($observed['Scan|WT1|10|17'][0] -eq 3 -and $observed['Scan|WT2|10|17'][0] -eq 1 -and $observed['Scan|WT3|10|17'][0] -eq 1) 'Scan Map expected3/1/1'
 foreach($phase in 1..3){Check ($observed["March|WT$phase|10|17"][0] -eq 1) 'March Map expected1/1/1';Check ($observed["Scan|WT$phase|11|17"][0] -eq 0) 'Hard Fail chip BLK255 not excluded';Check ($observed["Scan|WT$phase|11|18"][1] -eq 'No Data') 'DC chip must be gray'}
 Check ($observed['Scan|WT1|10|18'][0] -eq 0 -and $observed['Scan|WT2|10|18'][1] -eq 'No Data' -and $observed['Scan|WT3|10|18'][1] -eq 'No Data') 'WT2/3 no-test chip status wrong'
 $result=$app.Run("'"+$book.Name+"'!ImportFailAddresses.AuditFunctionBlocks",$book)
 Check ($result -eq 'Audit complete. Map BLKs=8; Complete=3; Partial=3; No Address=2; Excess=0') $result
 $report=$book.Worksheets.Item('Missing BLK Address');$actual=@{}
 for($r=5;$r -le 12;$r++){$v=for($c=1;$c -le 12;$c++){$report.Cells.Item($r,$c).Value2};$key=$v[0]+'|'+$v[1]+'|'+$v[5]+'|'+$v[6]+'|'+$v[7];$actual[$key]=@($v[8],$v[9],$v[10],$v[11])}
 foreach($expected in (Get-Content -LiteralPath (Join-Path $Project 'Sample_Raw_Data/expected_coverage.json') -Raw|ConvertFrom-Json)){
  $key=$expected[0]+'|'+$expected[1]+'|'+$expected[2]+'|'+$expected[3]+'|'+$expected[4]
  for($i=0;$i -lt 4;$i++){Check ($actual[$key][$i] -eq $expected[$i+5]) "Coverage truth mismatch:$key field$i"}
 }
 Check ([string]::IsNullOrEmpty($book.Worksheets.Item('Config').Range('H1:O38').Value2.GetValue(2,3))) 'Import wrote metadata behind moved buttons'
 $book.SaveCopyAs('C:/Users/dora9/outputs/ASC_ThreeASC_Project/wafer128/Home_Fixture_Validated_Result.xlsm')
 Write-Output 'PASS native home fixture: real file reads, unsorted STEP assignment, metadata, 1005 Address rows/183 columns,413 FBC rows, Scan3/1/1 and March1/1/1, gray no-test chips, Hard Fail, exact8 coverage rows with Complete3/Partial3/No Address2, and clean Config button area.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
