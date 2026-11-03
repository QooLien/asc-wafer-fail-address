param([string]$Source='C:/Users/dora9/outputs/_archive/ASC_2026-11-03_cleanup/user_layouts/ASC_ThreeASC_Log_Sheet.xlsm',[string]$Workbook='C:/Users/dora9/outputs/_archive/ASC_2026-11-03_cleanup/user_layouts/ASC_ThreeASC_Config_Updated.xlsm')
$ErrorActionPreference='Stop'
function Check($ok,$message){if(-not $ok){throw $message}}
$app=$null;$a=$null;$b=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $a=$app.Workbooks.Open($Source,0,$true);$b=$app.Workbooks.Open($Workbook,0,$true)
 $before=$a.Worksheets.Item('Config').Range('A1:G901').Value2;$cfg=$b.Worksheets.Item('Config');$after=$cfg.Range('A1:G901').Value2
 for($r=1;$r -le 901;$r++){for($c=1;$c -le 7;$c++){Check ($before.GetValue($r,$c) -eq $after.GetValue($r,$c)) "User setting changed at $r,$c"}}
 Check ($cfg.Range('G4').Value2 -eq 'Hard Fail' -and $cfg.Range('G20').Value2 -eq 'Scan' -and $cfg.Range('G38').Value2 -eq 'March') 'Stop Sign lost'
 foreach($sheet in $b.Worksheets){
  Check ($sheet.UsedRange.Font.Name -eq 'Calibri') ($sheet.Name+' font is not uniformly Calibri')
  Check ($sheet.UsedRange.Font.Size -eq 11) ($sheet.Name+' font size is not uniformly11')
  Check ($sheet.UsedRange.HorizontalAlignment -eq -4108) ($sheet.Name+' horizontal alignment differs')
  Check ($sheet.UsedRange.VerticalAlignment -eq -4108) ($sheet.Name+' vertical alignment differs')
 }
 $buttons=$cfg.Buttons();Check ($buttons.Count -eq 4) 'Expected four Forms buttons'
 $expected=@('ImportASCFailAddresses','ImportASCFBC','BuildFailMaps','FindMissingBLKAddresses');$actual=@()
 for($buttonIndex=1;$buttonIndex -le $buttons.Count;$buttonIndex++){
  $button=$buttons.Item($buttonIndex)
  Write-Output ('Button font: '+$button.Font.Name+' / '+$button.Font.Size)
  Check ($button.Font.Name -eq 'Calibri' -and $button.Font.Size -eq 11) 'Button font must be Calibri11'
  Check ($button.HorizontalAlignment -eq -4108 -and $button.VerticalAlignment -eq -4108) 'Button text must be centered'
  $actual+=($button.OnAction -replace '^.*!','')
 }
 foreach($macro in $expected){Check ($actual -contains $macro) ('Missing macro binding '+$macro)}
 $engine="'"+$b.Name+"'!BLKVminEngine."
 [void]$app.Run($engine+'BLK_Setup',$cfg)
 $tnos=$app.Run($engine+'BLK_AddressTNOs');Check ($tnos.Count -eq 59) 'Expected22 Monitor +37 Function TNOs'
 Check ([string]::IsNullOrEmpty($cfg.Range('H1:O38').Value2.GetValue(1,1))) 'Removed information restored unexpectedly'
 Write-Output 'PASS native user Config: all A:G settings preserved, 22 Monitor+37 Function TNOs valid, Hard Fail/Scan/March unchanged, four correct macro buttons, all sheet cells/buttons Calibri11 and centered.'
}finally{if($null -ne $b){$b.Close($false)};if($null -ne $a){$a.Close($false)};if($null -ne $app){$app.Quit()}}
