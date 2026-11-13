param([string]$Workbook=(Join-Path $PSScriptRoot '../templates/SRAM Extra BLK Failure  Analysis_v13.xlsm'),[string]$Output='C:/Users/dora9/outputs/Native_Address_Coverage_Test.xlsm')
$ErrorActionPreference='Stop'
Copy-Item -LiteralPath $Workbook -Destination $Output -Force
function Check($ok,$message){if(-not $ok){throw $message}}
function Run-Result($macro,$a,$b=$null){try{if($null -eq $b){return $app.Run($macro,$a)}else{return $app.Run($macro,$a,$b)}}catch [System.ArgumentNullException]{if($_.Exception.ParamName -eq 'ptr'){return ''};throw}}
function FBC-Line($tno,$values){$v=for($b=0;$b -lt 256;$b++){if($values.ContainsKey($b)){[string]$values[$b]}else{'0'}};return "Tno= $tno DUT=9 X=10 Y=17 FBC BY BLK= "+($v -join ' ')+"`n"}
function Compare-Grid($a,$b){for($r=1;$r -le $a.GetLength(0);$r++){for($c=1;$c -le $a.GetLength(1);$c++){Check ($a.GetValue($r,$c) -eq $b.GetValue($r,$c)) "Source modified $r,$c"}}}
$app=$null;$book=$null;$sourceBook=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Output,0,$false)
 $cfg=$book.Worksheets.Item('Config');$cfg.Range('D2:G901').ClearContents()
 foreach($s in @(@(2,500,''),@(3,501,'Scan'),@(4,490,'Hard Fail'))){$cfg.Cells.Item($s[0],4).Value2=[double]$s[1];$cfg.Cells.Item($s[0],5).Value2="Func$($s[1])";$cfg.Cells.Item($s[0],7).Value2=[string]$s[2]}
 $engine="'"+$book.Name+"'!BLKVminEngine."
 [void]$app.Run($engine+'BLK_Setup',$cfg)
 for($phase=1;$phase -le 3;$phase++){
  $endpoints=if($phase -eq 1){@{0=2;1=5;2=1;3=1;255=4}}elseif($phase -eq 2){@{0=2;4=1;5=3}}else{@{5=3;6=1}}
  $early=if($phase -eq 1){@{4=1}}else{@{}};$hard=if($phase -eq 1){@{3=1}}else{@{}}
  $raw=(FBC-Line 500 $early)+(FBC-Line 501 $endpoints)+(FBC-Line 490 $hard)+(FBC-Line 350 @{7=99})
  [void]$app.Run($engine+'BLK_BeginFile',"WT$phase")
  [void]$app.Run($engine+'BLK_ObserveLocations',$raw,$true);[void]$app.Run($engine+'BLK_Feed',$raw,$true)
  $msg=Run-Result ($engine+'BLK_WritePhase') $book $phase;Check ([string]::IsNullOrEmpty($msg)) $msg
 }
 [void](Run-Result ($engine+'BLK_Finalize') $book);[void](Run-Result ($engine+'BLK_BuildFailMaps') $book)
 $counts=$book.Worksheets.Item('BLK Fail Count');Check ($counts.Range('H2').Value2 -eq 4 -and $counts.Range('H3').Value2 -eq 1 -and $counts.Range('H4').Value2 -eq 1) 'Map baseline must count 4/1/1 BLKs'
 $address=$book.Worksheets.Item('Fail Address');$address.UsedRange.ClearContents();$headers=New-Object 'object[,]' 1,15
 for($c=0;$c -lt 6;$c++){$headers[0,$c]=@('X','Y','BLK','WL','BL','IO')[$c]}
 for($phase=1;$phase -le 3;$phase++){for($j=0;$j -lt 3;$j++){$headers[0,(6+($phase-1)*3+$j)]="WT$phase`n$(@(500,501,490)[$j])"}}
 $address.Range('A1:O1').Value2=$headers;$detail=New-Object 'object[,]' 11,15
 $entries=@(@(0,0,7,10),@(0,1,7,10),@(1,0,7,10),@(1,1,7,10),@(2,0,7,10),@(2,1,7,10),@(5,0,10,10),@(6,0,13,10),@(255,0,6,10),@(255,1,10,10),@(255,2,7,11))
 for($r=0;$r -lt $entries.Count;$r++){$e=$entries[$r];$detail[$r,0]=[double]$e[3];$detail[$r,1]=[double]17;$detail[$r,2]=[double]$e[0];$detail[$r,3]=[double]$e[1];$detail[$r,4]=[double]0;$detail[$r,5]=[double]0;$detail[$r,$e[2]]=[double]1}
 $address.Range('A2:O12').Value2=$detail;$fbc=$book.Worksheets.Item('FBC by BLK');$addrBefore=$address.Range('A1:O12').Value2;$fbcBefore=$fbc.UsedRange.Value2
 $macro="'"+$book.Name+"'!ImportFailAddresses.AuditFunctionBlocks";$expected='*Map BLKs=6; Complete=2; Partial=2; No Address=1; Excess=1'
 $result=$app.Run($macro,$book);Check ($result -like $expected) $result
 $report=$book.Worksheets.Item('Missing BLK Address');$found=@{}
 for($r=5;$r -le 10;$r++){$key=[string]$report.Cells.Item($r,1).Value2+'|'+$report.Cells.Item($r,2).Value2+'|'+$report.Cells.Item($r,8).Value2;$found[$key]=@($report.Cells.Item($r,10).Value2,$report.Cells.Item($r,11).Value2,$report.Cells.Item($r,12).Value2)}
 Check ($found['WT1|501|0'][2] -eq 'Complete' -and $found['WT1|501|0'][0] -eq 2) 'Complete count wrong'
 Check ($found['WT1|501|1'][2] -eq 'Partial Address' -and $found['WT1|501|1'][1] -eq 3) 'Partial gap wrong'
 Check ($found['WT1|501|255'][2] -eq 'No Address' -and $found['WT1|501|255'][0] -eq 0) 'Wrong WT/TNO/Chip contaminated BLK255'
 Check ($found['WT1|501|2'][2] -eq 'Address Count > FBC') 'Excess was classified as complete'
 Check (-not $found.ContainsKey('WT1|501|3') -and -not $found.ContainsKey('WT2|501|4') -and -not $found.ContainsKey('WT3|501|5')) 'Excluded BLKs leaked into coverage'
 Compare-Grid $addrBefore ($address.Range('A1:O12').Value2);Compare-Grid $fbcBefore ($fbc.UsedRange.Value2)
 $book.SaveCopyAs(($Output -replace '.xlsm$','_Valid.xlsm'))
 [void]$address.Range('A1:O12').AutoFilter(3,0);[void]$fbc.UsedRange.AutoFilter(2,350);$cfg.Range('Q2').Value2='All Function'
 $result=$app.Run($macro,$book);Check ($result -like $expected) 'Filters or old Q2 changed scope';$address.AutoFilterMode=$false;$fbc.AutoFilterMode=$false
 $target=0;for($r=2;$r -le $fbcBefore.GetLength(0);$r++){if($fbc.Cells.Item($r,1).Value2 -eq 'WT1' -and $fbc.Cells.Item($r,2).Value2 -eq 501){$target=$r}}
 $fbc.Cells.Item($target,260).Value2=[double]2147483648;$result=$app.Run($macro,$book);$report=$book.Worksheets.Item('Missing BLK Address')
 $large=$false;for($r=5;$r -le 10;$r++){if($report.Cells.Item($r,1).Value2 -eq 'WT1' -and $report.Cells.Item($r,8).Value2 -eq 255){$large=($report.Cells.Item($r,11).Value2 -eq 2147483648)}};Check $large 'Large gap overflow';$fbc.Cells.Item($target,260).Value2=[double]4
 $savedTitle=$report.Range('A1').Value2;$index=$book.Worksheets.Item('Import Index');$index.Range('H2').Value2='Address Pending'
 $result=$app.Run($macro,$book);Check ($result -like '*incomplete*') 'Pending import accepted';$index.Range('H2').ClearContents()
 $book.Worksheets.Item('Chip Locations').Range('E3').Value2='Data Error';$result=$app.Run($macro,$book);Check ($result -like '*parse/output error*') 'Parse error accepted';$book.Worksheets.Item('Chip Locations').Range('E3').Value2='OK'
 $last=$fbcBefore.GetLength(0)+1;$fbc.Cells.Item($last,1).Resize(1,260).Value2=$fbc.Cells.Item($target,1).Resize(1,260).Value2
 $result=$app.Run($macro,$book);Check ($result -like '*Duplicate FBC*') 'Duplicate FBC accepted';$fbc.Cells.Item($last,1).Resize(1,260).ClearContents()
 Check ($book.Worksheets.Item('Missing BLK Address').Range('A1').Value2 -eq $savedTitle) 'Failure replaced report';Check ($book.Worksheets.Count -eq 10) 'Staging sheet leaked'
 $result=$app.Run($macro,$book);Check ($result -like $expected) $result
 $sourcePath=$Output -replace '.xlsm$','_Source.xlsm';$book.SaveCopyAs($sourcePath);$sourceBook=$app.Workbooks.Open($sourcePath,0,$true)
 $result=$app.Run($macro,$sourceBook,$book);Check ($result -like $expected) 'Read-only input failed';Check $sourceBook.ReadOnly 'Source should be read-only';$sourceBook.Close($false);$sourceBook=$null
 Compare-Grid $addrBefore ($address.Range('A1:O12').Value2);Compare-Grid $fbcBefore ($fbc.UsedRange.Value2)
 Write-Output 'PASS native Map Address Coverage: Map 4/1/1, Complete/Partial/No Address/Excess counts, exact WT/TNO/Chip/BLK, BLK255, large gap, Hard Fail/prior all-Function exclusion, hidden rows, source/error retention, read-only input and unchanged sources.'
}finally{if($null -ne $sourceBook){$sourceBook.Close($false)};if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
