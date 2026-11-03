param([string]$Workbook='C:/Users/dora9/outputs/ASC_ThreeASC_Soft_Fail_Map.xlsm',[string]$Output='C:/Users/dora9/outputs/Native_Soft_Map_Test.xlsm')
$ErrorActionPreference='Stop'
Copy-Item -LiteralPath $Workbook -Destination $Output -Force
function Check($ok,$message){if(-not $ok){throw $message}}
function Run-Result($macro,$arg1,$arg2=$null){
 try{if($null -eq $arg2){return $app.Run($macro,$arg1)}else{return $app.Run($macro,$arg1,$arg2)}}
 catch [System.ArgumentNullException]{if($_.Exception.ParamName -eq 'ptr'){return ''};throw}
}
function Line($tno,$blocks){
 $values=for($b=0;$b -lt 256;$b++){if($blocks -contains $b){'5'}else{'0'}}
 return "Tno= $tno DUT= 9 X= 10 Y= 17 FBC BY BLK= "+($values -join ' ')+"`n"
}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Output,0,$false)
 $cfg=$book.Worksheets.Item('Config');$cfg.Range('D2:G901').ClearContents()
 foreach($setting in @(@(2,490,'Hard Fail'),@(3,491,'hard fail'),@(4,501,'Scan'),@(5,550,''),@(6,599,''))){
  $r=$setting[0];$cfg.Cells.Item($r,4).Value2=[double]$setting[1];$cfg.Cells.Item($r,5).Value2='Function';$cfg.Cells.Item($r,7).Value2=[string]$setting[2]
 }
 $address=$book.Worksheets.Item('Fail Address');$address.Range('A2').Value2=[double]10;$address.Range('G2').Value2=[double]1
 $prefix="'"+$book.Name+"'!BLKVminEngine."
 [void]$app.Run($prefix+'BLK_Setup',$cfg)
 for($phase=1;$phase -le 3;$phase++){
  [void]$app.Run($prefix+'BLK_BeginFile',"test-WT$phase")
  $target=if($phase -eq 1){@(0,1,255)}elseif($phase -eq 2){@(0,1,2,3,255)}else{@(0,1,2,3,4,255)}
  $early=if($phase -eq 1){@(3)}else{@()}
  $raw=(Line 490 @(0))+(Line 491 @(255))+(Line 501 $target)+(Line 550 $early)+(Line 599 @())
  [void]$app.Run($prefix+'BLK_ObserveLocations',$raw,$true);[void]$app.Run($prefix+'BLK_Feed',$raw,$true)
  $msg=Run-Result ($prefix+'BLK_WritePhase') $book $phase;Check ([string]::IsNullOrEmpty($msg)) $msg
 }
 [void](Run-Result ($prefix+'BLK_Finalize') $book)
 $fbc=$book.Worksheets.Item('FBC by BLK');$before=$fbc.UsedRange.Value2
 $msg=Run-Result ($prefix+'BLK_BuildFailMaps') $book;Check ([string]::IsNullOrEmpty($msg)) $msg
 $summary=$book.Worksheets.Item('BLK Fail Count')
 Check ($summary.Range('H2').Value2 -eq 1) 'WT1: Hard Fail blocks0/255 must be excluded'
 Check ($summary.Range('H3').Value2 -eq 1) 'WT2: early Function550 block3 and Pattern501 block1 must be excluded even when final599 is zero'
 Check ($summary.Range('H4').Value2 -eq 1) 'WT3: all earlier Function failures must be excluded; only block4 remains'
 $names=@($book.Worksheets|ForEach-Object{$_.Name});Check ($names -notcontains 'Hard Fail Fail Map') 'Hard Fail marker produced a pattern map'
 $after=$fbc.UsedRange.Value2
 for($r=1;$r -le $before.GetLength(0);$r++){for($c=1;$c -le 260;$c++){Check ($before.GetValue($r,$c) -eq $after.GetValue($r,$c)) "FBC source changed $r,$c"}}
 Check ($address.Range('G2').Value2 -eq 1) 'Address source changed'
 $book.SaveCopyAs(($Output -replace '.xlsm$','_Valid.xlsm'))
 # A newly hard-failing BLK in WT2 remains excluded in WT3 even if WT3 hard record is zero.
 $rowWT2Hard=0;$rowWT1Hard=0
 for($r=2;$r -le $after.GetLength(0);$r++){
  if($fbc.Cells.Item($r,2).Value2 -eq 490){
   if($fbc.Cells.Item($r,1).Value2 -eq 'WT2'){$rowWT2Hard=$r}
   if($fbc.Cells.Item($r,1).Value2 -eq 'WT1'){$rowWT1Hard=$r}
  }
 }
 $fbc.Cells.Item($rowWT2Hard,9).Value2=[double]1
 [void](Run-Result ($prefix+'BLK_BuildFailMaps') $book)
 Check ($summary.Range('H3').Value2 -eq 1 -and $summary.Range('H4').Value2 -eq 0) 'Hard Fail must persist across later WT'
 # Required missing Hard Fail record is unavailable, never treated as Pass.
 $fbc.Cells.Item($rowWT1Hard,2).Value2=[double]492
 [void](Run-Result ($prefix+'BLK_BuildFailMaps') $book)
 Check ($summary.Range('I2').Value2 -eq 'No Data' -and $summary.Range('I4').Value2 -eq 'No Data') 'Missing prerequisite must not become zero/Pass'
 # Without Hard Fail markers, prior-WT all Function filtering remains active.
 $fbc.Cells.Item($rowWT1Hard,2).Value2=[double]490;$cfg.Range('G2:G3').ClearContents()
 $fbc.Cells.Item($rowWT2Hard,9).Value2=[double]0
 [void](Run-Result ($prefix+'BLK_BuildFailMaps') $book)
 Check ($summary.Range('H2').Value2 -eq 3 -and $summary.Range('H3').Value2 -eq 1 -and $summary.Range('H4').Value2 -eq 1) 'Prior-WT all Function filter without Hard Fail failed'
 Write-Output 'PASS native Excel: per-BLK Hard Fail union, multiple case-insensitive markers, WT1/2 ALL Function gates, BLK0/255, permanent exclusion, missing prerequisites, no-marker mode, and unchanged FBC/Address sources.'
}finally{
 if($null -ne $book){$book.Close($false)}
 if($null -ne $app){$app.Quit()}
}
