param([Parameter(Mandatory=$true)][string]$Workbook,[string]$Output='C:/Users/dora9/outputs/ASC_ThreeASC_Project/wafer128/Native_Pattern_Sequence.xlsm')
$ErrorActionPreference='Stop';Copy-Item -LiteralPath $Workbook -Destination $Output -Force
function Check($ok,$message){if(-not $ok){throw $message}}
function Run-Result($macro,$a,$b=$null){try{if($null -eq $b){return $app.Run($macro,$a)}else{return $app.Run($macro,$a,$b)}}catch [System.ArgumentNullException]{if($_.Exception.ParamName -eq 'ptr'){return ''};throw}}
function Raw-Line($tno,$blocks){$v=for($b=0;$b -lt 256;$b++){if($blocks -contains $b){'10'}else{'0'}};return "Tno= $tno DUT=9 X=10 Y=17 FBC BY BLK= "+($v -join ' ')+"`n"}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Output,0,$false);$cfg=$book.Worksheets.Item('Config');$cfg.Range('A2:G901').ClearContents()
 $cfg.Range('A2').Value2=[double]350;$cfg.Range('B2').Value2='Monitor'
 foreach($setting in @(@(2,502,'Hard Fail'),@(3,530,'Scan'),@(4,717,'March'),@(5,650,''),@(6,510,'Post'))){$r=$setting[0];$cfg.Cells.Item($r,4).Value2=[double]$setting[1];$cfg.Cells.Item($r,5).Value2="Function$($setting[1])";$cfg.Cells.Item($r,7).Value2=[string]$setting[2]}
 $prefix="'"+$book.Name+"'!BLKVminEngine."
 [void]$app.Run($prefix+'BLK_Setup',$cfg)
 for($phase=1;$phase -le 3;$phase++){
  if($phase -eq 1){$hard=@(0);$scan=@(0,1,255);$march=@(0,1,2,255);$post=@(0,1,2,3,255);$extra=@(4)}
  elseif($phase -eq 2){$hard=@();$scan=@(1,2,4,5);$march=@(1,2,4,5,6);$post=@(1,2,4,5,6,7);$extra=@()}
  else{$hard=@();$scan=@(5,6,7,8);$march=@(5,6,7,8,9);$post=@(5,6,7,8,9,10);$extra=@()}
  $raw=(Raw-Line 502 $hard)+(Raw-Line 530 $scan)+(Raw-Line 717 $march)+(Raw-Line 650 $extra)+(Raw-Line 510 $post)+(Raw-Line 350 @(2,5,8))
  [void]$app.Run($prefix+'BLK_BeginFile',"WT$phase")
  [void]$app.Run($prefix+'BLK_ObserveLocations',$raw,$true);[void]$app.Run($prefix+'BLK_Feed',$raw,$true)
  $msg=Run-Result ($prefix+'BLK_WritePhase') $book $phase;Check ([string]::IsNullOrEmpty($msg)) $msg
 }
 [void](Run-Result ($prefix+'BLK_Finalize') $book);$fbc=$book.Worksheets.Item('FBC by BLK');$before=$fbc.UsedRange.Value2
 [void](Run-Result ($prefix+'BLK_BuildFailMaps') $book)
 $summary=$book.Worksheets.Item('BLK Fail Count');$counts=@{}
 for($r=2;$r -le 10;$r++){$counts[([string]$summary.Cells.Item($r,1).Value2+'|'+$summary.Cells.Item($r,5).Value2)]=$summary.Cells.Item($r,8).Value2}
 Check ($counts['Scan|WT1'] -eq 2) 'Hard Fail must be removed before Scan'
 foreach($phase in 1..3){Check ($counts["March|WT$phase"] -eq 1) "WT$phase March included prior Scan Fail";Check ($counts["Post|WT$phase"] -eq 1) "WT$phase Post included earlier pattern Fail"}
 Check ($counts['Scan|WT2'] -eq 1 -and $counts['Scan|WT3'] -eq 1) 'Previous WT all-Function gate broken'
 $masks=$app.Run($prefix+'BLK_MapIncludedMasks',$book)
 $mask=[string]$masks.Item('WT1|717|10|17');Check ($mask.Substring(2,1) -eq '1' -and $mask.Substring(1,1) -eq '0' -and $mask.Substring(255,1) -eq '0') 'Coverage scope differs from March Map'
 $mask=[string]$masks.Item('WT2|717|10|17');Check ($mask.Substring(6,1) -eq '1' -and $mask.Substring(5,1) -eq '0') 'WT2 same-stage Scan gate absent'
 $after=$fbc.UsedRange.Value2;for($r=1;$r -le $before.GetLength(0);$r++){for($c=1;$c -le 260;$c++){Check ($before.GetValue($r,$c) -eq $after.GetValue($r,$c)) 'FBC data altered'}}
 # No Address has been imported: Map must still work solely from FBC.
 Check ([string]::IsNullOrEmpty($book.Worksheets.Item('Fail Address').Range('A2').Value2)) 'Fixture expected no Address'
 $book.SaveCopyAs(($Output -replace '.xlsm$','_Valid.xlsm'))
 # Missing earlier Scan cannot be accepted as Pass for March.
 $row=0;for($r=2;$r -le $after.GetLength(0);$r++){if($fbc.Cells.Item($r,1).Value2 -eq 'WT1' -and $fbc.Cells.Item($r,2).Value2 -eq 530){$row=$r}}
 $fbc.Cells.Item($row,2).Value2=[double]531
 [void](Run-Result ($prefix+'BLK_BuildFailMaps') $book)
 Check ($summary.Range('I5').Value2 -eq 'No Data') 'Missing Scan was treated as Pass for March'
 Write-Output 'PASS sequence: Hard Fail first; Scan->March->Post within each WT using Config order (not numeric TNO); prior WT ALL Function Pass; BLK0/255; Monitor excluded; no Address dependency; shared coverage scope; missing predecessor unavailable; raw FBC unchanged.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
