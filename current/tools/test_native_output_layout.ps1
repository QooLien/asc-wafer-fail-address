param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false);$cfg=$book.Worksheets.Item('Config');$buttons=$cfg.Buttons();$positions=@()
 $configHeight=$cfg.Rows('1:901').RowHeight
 for($i=1;$i -le $buttons.Count;$i++){$b=$buttons.Item($i);$positions+=,@($b.Left,$b.Top,$b.Width,$b.Height)}
 [void]$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_ApplyPresentation",$book)
 $view=$book.Worksheets.Item('BLK_FBC_Comparison');$mapping=$book.Worksheets.Item('Import Index').Range('Q2:U178').Value2
 $functionCount=0;$monitorCount=0
 for($r=1;$r -le 177;$r++){
  $c=[int]$mapping.GetValue($r,3);$head=$view.Cells.Item(2,$c)
  if($mapping.GetValue($r,4) -eq 'Function'){
   Check ($head.Orientation -in @(90,-4171)) ('Function heading is not bottom-to-top:'+ $head.Orientation);$functionCount++
   Check ($view.Columns.Item($c).ColumnWidth -ge 7 -and $view.Columns.Item($c).ColumnWidth -le 14) 'Function column sizing invalid'
  }else{Check ($head.Orientation -in @(90,-4171)) ('Monitor heading is not bottom-to-top:'+ $head.Orientation);$monitorCount++}
 }
 Check ($functionCount -eq 111 -and $monitorCount -eq 66) 'Group counts changed'
 Check ($view.Rows.Item(2).RowHeight -ge 90 -and $view.Rows.Item(2).RowHeight -le 200) 'Rotated header row sizing invalid'
 if($view.UsedRange.Rows.Count -gt 2){Check ($view.Range('A3:FX1026').RowHeight -eq 26) 'Comparison data row heights differ'}
 $expected=@('匯入 Fail Address','核對 BLK Address','更新 Map','匯入 FBC')
 # Native button order is kept from the user's workbook; evaluate captions by macro action.
 $captions=@{}
 for($i=1;$i -le $buttons.Count;$i++){
  $b=$buttons.Item($i);$captions[($b.OnAction -replace '^.*!','')]=$b.Caption
  Check ($b.Font.Name -eq 'Calibri' -and $b.Font.Size -eq 11) 'Button font changed'
  Check ($b.Left -eq $positions[$i-1][0] -and $b.Top -eq $positions[$i-1][1] -and $b.Width -eq $positions[$i-1][2] -and $b.Height -eq $positions[$i-1][3]) 'Button position or size changed'
 }
 Check ($captions['ImportASCFailAddresses'] -eq $expected[0] -and $captions['FindMissingBLKAddresses'] -eq $expected[1] -and $captions['BuildFailMaps'] -eq $expected[2] -and $captions['ImportASCFBC'] -eq $expected[3]) 'Button caption is not concise'
 Check ($cfg.Rows('1:901').RowHeight -eq $configHeight) 'Config row heights changed'
 $book.Save()
 Write-Output 'PASS native layout:ALL177 Monitor/Function TNO headings rotate90 bottom-to-top; bounded automatic widths/header heights; uniform26pt data rows; simple Calibri11 buttons with unchanged bounds; Config heights retained.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
