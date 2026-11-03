param([Parameter(Mandatory=$true)][string]$Workbook,[string]$Output='')
$ErrorActionPreference='Stop'
function Check($ok,$message){if(-not $ok){throw $message}}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.ScreenUpdating=$false
 $book=$app.Workbooks.Open($Workbook,0,$false);$raw=$book.Worksheets.Item('FBC by BLK');$before=$raw.UsedRange.Value2
 try{$result=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_BuildComparison",$book)}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw};$result=''}
 Check ([string]::IsNullOrEmpty($result)) $result
 $view=$book.Worksheets.Item('BLK_FBC_Comparison');$grid=$view.UsedRange.Value2
 Check ($grid.GetLength(0) -eq 1026 -and $grid.GetLength(1) -eq 180) 'Expected4 Chips x256 BLKs and3x59 TNOs'
 Check ($view.Range('A1').Value2 -eq 'Chip / BLK') 'Chip heading moved'
 foreach($heading in @(@('D1','WT1 / Monitor',22),@('Z1','WT2 / Monitor',22),@('AV1','WT3 / Monitor',22),@('BR1','WT1 / Function',37),@('DC1','WT2 / Function',37),@('EN1','WT3 / Function',37))){
  Check ($view.Range($heading[0]).Value2 -eq $heading[1]) ('Group heading moved: '+$heading[1])
  Check ($view.Range($heading[0]).MergeArea.Columns.Count -eq $heading[2]) 'Group width wrong'
 }
 $mapping=$book.Worksheets.Item('Import Index').Range('Q2:U178').Value2;$columns=@{};$monitors=@{}
 for($r=1;$r -le 177;$r++){
  $key=[string]$mapping.GetValue($r,1)+'|'+$mapping.GetValue($r,2);$col=[int]$mapping.GetValue($r,3);$columns[$key]=$col
  Check ($grid.GetValue(2,$col) -eq $mapping.GetValue($r,5)) 'Header label differs from numeric mapping'
  if($mapping.GetValue($r,4) -eq 'Monitor'){$wt=[string]$mapping.GetValue($r,1);$monitors[$wt]++}
 }
 foreach($wt in @('WT1','WT2','WT3')){Check ($monitors[$wt] -eq 22) "$wt is missing Monitor columns"}
 $rows=@{}
 for($r=3;$r -le 1026;$r++){$key=[string]$grid.GetValue($r,1)+'|'+$grid.GetValue($r,2)+'|'+$grid.GetValue($r,3);$rows[$key]=$r}
 $checked=0
 for($r=2;$r -le $before.GetLength(0);$r++){
  $col=$columns[([string]$before.GetValue($r,1)+'|'+$before.GetValue($r,2))]
  Check ($null -ne $col) 'Raw TNO missing from comparison'
  for($b=0;$b -lt 256;$b++){
   $key=[string]$before.GetValue($r,3)+'|'+$before.GetValue($r,4)+'|'+$b;$destRow=$rows[$key]
   Check ($grid.GetValue($destRow,$col) -eq $before.GetValue($r,($b+5))) "FBC index mismatch source row$r BLK$b"
   $checked++
  }
 }
 $after=$raw.UsedRange.Value2
 for($r=1;$r -le $before.GetLength(0);$r++){for($c=1;$c -le 260;$c++){Check ($before.GetValue($r,$c) -eq $after.GetValue($r,$c)) 'Internal raw values changed'}}
 Check ($raw.Visible -eq 2) 'Internal FBC table must be very hidden'
 Check ($view.UsedRange.Font.Name -eq 'Calibri' -and $view.UsedRange.Font.Size -eq 11) 'Comparison font differs'
 Check ($view.UsedRange.HorizontalAlignment -eq -4108 -and $view.UsedRange.VerticalAlignment -eq -4108) 'Comparison must be centered'
 Check (-not $app.ActiveWindow.DisplayGridlines) 'Comparison gridlines should be off'
 $missingRow=$rows['11|18|0'];$col=$columns['WT2|350'];Check ([string]::IsNullOrEmpty($view.Cells.Item($missingRow,$col).Value2)) 'Missing record became numeric zero'
 Check ($view.Cells.Item($missingRow,$col).Interior.Color -eq 14737632) 'Missing data should be gray'
 $zeroRow=$rows['10|17|0'];$col=$columns['WT2|350'];Check ($view.Cells.Item($zeroRow,$col).Value2 -eq 0 -and $view.Cells.Item($zeroRow,$col).Text -eq '') 'Real zero should display blank'
 foreach($wt in @('WT1','WT2','WT3')){Check ($grid.GetValue($rows['10|17|9'],$columns["$wt|350"]) -eq 9) "$wt Monitor FBC lost"}
 Check ($view.Cells.Item(2,$columns['WT3|351']).Value2 -eq 'W-Low pBL_1.035v') 'Three-decimal voltage lost'
 if($Output){$book.SaveAs($Output,52)}else{$book.Save()}
 Write-Output "PASS native FBC comparison: $checked FBC values match, BLK0/255 orientation,4 Chips,59 TNOs per WT,22 Monitor columns in EACH WT, Func/Voltage headers, blank zeros vs gray missing, numeric TNO mapping, hidden raw unchanged and Calibri11 centered."
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
