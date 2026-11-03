param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
function Click-Map($name,$address){
 [void]$app.Run("'"+$book.Name+"'!ThisWorkbook.FixtureDoubleClick",$name,$address)
 $handled=$book.Worksheets.Item('Import Index').Range('AL1').Value2
 try{$msg=$app.Run("'"+$book.Name+"'!ImportFailAddresses.FixtureLastMessage")}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw};$msg=''}
 Check ([string]::IsNullOrEmpty($msg)) $msg
 return $handled
}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$true
 $book=$app.Workbooks.Open($Workbook,0,$false)
 try{$msg=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_BuildFailMaps",$book)}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw};$msg=''}
 Check ([string]::IsNullOrEmpty($msg)) $msg
 $raw=$book.Worksheets.Item('FBC by BLK');$address=$book.Worksheets.Item('Fail Address');$beforeRaw=$raw.UsedRange.Value2;$beforeAddress=$address.UsedRange.Value2
 Check (-not (Click-Map 'Scan Fail Map' 'B4')) 'Axis should not drill down'
 Check (-not (Click-Map 'Scan Fail Map' 'C5')) 'Zero should not drill down'
 Check (-not (Click-Map 'Scan Fail Map' 'C6')) 'No Data should not drill down'
 $handled=Click-Map 'Scan Fail Map' 'B5'
 Check $handled 'Positive Scan should cancel editing and open details'
 $detail=$book.Worksheets.Item('Chip Fail Details');Check ($book.ActiveSheet.Name -eq 'Chip Fail Details') 'Event did not navigate'
 Check ($detail.Range('A1').Value2 -like 'Scan | WT1 | X=10 Y=17*') 'Chip context wrong'
 Check ($detail.Range('C6').Value2 -eq 530 -and $detail.Range('H6').Value2 -eq 0 -and $detail.Range('I6').Value2 -eq 1005 -and $detail.Range('J6').Value2 -eq 996 -and $detail.Range('K6').Value2 -eq 9) 'Scan BLK0 count/gap wrong'
 Check ($detail.Range('H7').Value2 -eq 1 -and $detail.Range('J7').Value2 -eq 2 -and $detail.Range('L7').Value2 -eq 'Partial Address') 'Partial Scan BLK1 wrong'
 Check ($detail.Range('H8').Value2 -eq 2 -and $detail.Range('J8').Value2 -eq 0 -and $detail.Range('L8').Value2 -eq 'No Address') 'Missing Scan BLK2 wrong'
 Check ($detail.Range('M6').Hyperlinks.Item(1).SubAddress -eq "'Chip Fail Details'!A12") 'First Address link wrong'
 Check ($detail.Range('M7').Hyperlinks.Item(1).SubAddress -eq "'Chip Fail Details'!A1008") 'Second Address link wrong'
 Check ($detail.Range('A2').Hyperlinks.Item(1).SubAddress -eq "'Scan Fail Map'!B5") 'Back link wrong'
 $addresses=$detail.Range('A12:K1009').Value2
 for($r=1;$r -le 998;$r++){
  Check ($addresses.GetValue($r,1) -eq 'WT1' -and $addresses.GetValue($r,2) -eq 530 -and $addresses.GetValue($r,5) -eq 10 -and $addresses.GetValue($r,6) -eq 17) 'Address mixed WT/TNO/Chip'
  $blk=if($r -le 996){0}else{1};$wl=if($r -le 996){$r-1}else{$r-997}
  Check ($addresses.GetValue($r,7) -eq $blk -and $addresses.GetValue($r,8) -eq $wl -and $addresses.GetValue($r,9) -eq 0 -and $addresses.GetValue($r,10) -eq 0) 'Address BLK/WL/BL/IO shifted'
 }
 Check ($detail.UsedRange.Font.Name -eq 'Calibri' -and $detail.UsedRange.Font.Size -eq 11) 'Details font wrong'
 Check (-not $app.ActiveWindow.DisplayGridlines) 'Details gridlines should be hidden'
 $book.SaveCopyAs('C:/Users/dora9/outputs/ASC_ThreeASC_Project/wafer128/Native_Drill_Scan_Valid.xlsm')
 Check (Click-Map 'March Fail Map' 'B5') 'March event failed'
 $detail=$book.Worksheets.Item('Chip Fail Details')
 Check ($detail.Range('C6').Value2 -eq 717 -and $detail.Range('H6').Value2 -eq 6 -and $detail.Range('J6').Value2 -eq 2 -and $detail.Range('L6').Value2 -eq 'Complete') 'March scope/count wrong'
 Check ($detail.Range('C7').Value2 -eq $null) 'Scan-failing BLK leaked into March details'
 Check (Click-Map 'March Fail Map' 'F5') 'WT2 March event failed'
 $detail=$book.Worksheets.Item('Chip Fail Details');Check ($detail.Range('A6').Value2 -eq 'WT2' -and $detail.Range('H6').Value2 -eq 7 -and $detail.Range('J6').Value2 -eq 0 -and $detail.Range('K6').Value2 -eq 4 -and $detail.Range('M6').Hyperlinks.Count -eq 0) 'WT2 missing Address should not borrow another WT'
 Check (Click-Map 'Scan Fail Map' 'J5') 'WT3 Scan event failed'
 $detail=$book.Worksheets.Item('Chip Fail Details');Check ($detail.Range('A6').Value2 -eq 'WT3' -and $detail.Range('C6').Value2 -eq 530 -and $detail.Range('H6').Value2 -eq 5 -and $detail.Range('J6').Value2 -eq 1 -and $detail.Range('K6').Value2 -eq 2) 'WT3 exact address key wrong'
 $afterRaw=$raw.UsedRange.Value2;for($r=1;$r -le $beforeRaw.GetLength(0);$r++){for($c=1;$c -le 260;$c++){Check ($beforeRaw.GetValue($r,$c) -eq $afterRaw.GetValue($r,$c)) 'Raw FBC changed'}}
 $afterAddress=$address.UsedRange.Value2;for($r=1;$r -le $beforeAddress.GetLength(0);$r++){for($c=1;$c -le $beforeAddress.GetLength(1);$c++){Check ($beforeAddress.GetValue($r,$c) -eq $afterAddress.GetValue($r,$c)) 'Source Address changed'}}
 # Explicit FBC100 /50 captured addresses verifies output-cap interpretation.
 $index=$book.Worksheets.Item('Import Index').Range('A2:D178').Value2;$flagCol=0
 for($r=1;$r -le 177;$r++){if($index.GetValue($r,1) -eq 'WT1' -and $index.GetValue($r,2) -eq '717'){$flagCol=[int]$index.GetValue($r,3)}}
 $extra=New-Object 'object[,]' 48,183
 for($r=0;$r -lt 48;$r++){$extra[$r,0]=[double]10;$extra[$r,1]=[double]17;$extra[$r,2]=[double]6;$extra[$r,3]=[double]($r+2);$extra[$r,4]=[double]0;$extra[$r,5]=[double]0;$extra[$r,($flagCol-1)]=[double]1}
 $address.Cells.Item(1007,1).Resize(48,183).Value2=$extra
 for($r=2;$r -le 414;$r++){if($raw.Cells.Item($r,1).Value2 -eq 'WT1' -and $raw.Cells.Item($r,2).Value2 -eq 717 -and $raw.Cells.Item($r,3).Value2 -eq 10 -and $raw.Cells.Item($r,4).Value2 -eq 17){$raw.Cells.Item($r,11).Value2=[double]100}}
 Check (Click-Map 'March Fail Map' 'B5') 'Output-cap case failed'
 $detail=$book.Worksheets.Item('Chip Fail Details');Check ($detail.Range('I6').Value2 -eq 100 -and $detail.Range('J6').Value2 -eq 50 -and $detail.Range('K6').Value2 -eq 50 -and $detail.Range('L6').Value2 -eq 'Partial Address') 'FBC100/50 Address gap wrong'
 Check ($book.Worksheets.Item($book.Worksheets.Count).Name -eq 'Log' -and $book.Worksheets.Item(($book.Worksheets.Count-1)).Name -eq 'Chip Locations') 'Last sheet order changed'
 Write-Output 'PASS workbook event drilldown: positive double-click cancels editing/navigates; axes/zero/missing ignored; Pattern scope and sequential BLKs correct;998 exact Scan addresses; complete/partial/missing and FBC100/Address50; correct WT/TNO/Chip/BLK; working internal/back links; source unchanged; Calibri11 and final sheet order.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
