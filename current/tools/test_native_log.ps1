param([string]$Workbook=(Join-Path $PSScriptRoot '../templates/SRAM Extra BLK Failure  Analysis_v13.1.xlsm'),[string]$Output='C:/Users/dora9/outputs/Native_Log_Test.xlsm')
$ErrorActionPreference='Stop';Copy-Item -LiteralPath $Workbook -Destination $Output -Force
function Check($ok,$message){if(-not $ok){throw $message}}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Output,0,$false);$cfg=$book.Worksheets.Item('Config');$log=$book.Worksheets.Item('Log')
 $cfg.Range('G2').Value2='Hard Fail';$cfg.Range('G3').Value2='Scan';$before=$cfg.Range('A1:G23').Value2
 Check ($cfg.Rows('1:901').RowHeight -eq 30) 'Config template row heights differ'
 $macro="'"+$book.Name+"'!ImportFailAddresses.WriteOperationLog"
 foreach($operation in @('Map','FBC','Address','Address Coverage')){[void]$app.Run($macro,$book,$operation,('Detail line'+"`n")*200,$true)}
 Check ($cfg.Rows('1:901').RowHeight -eq 30) 'Long log changed Config row heights'
 Check ([string]::IsNullOrEmpty($cfg.Range('T2').Value2)) 'Config still contains long messages'
 for($r=1;$r -le 23;$r++){for($c=1;$c -le 7;$c++){Check ($before.GetValue($r,$c) -eq $cfg.Cells.Item($r,$c).Value2) 'Function/Stop Sign configuration changed'}}
 for($r=2;$r -le 5;$r++){Check ($log.Cells.Item($r,3).Value2 -eq 'Warning / Error') 'Log status wrong';Check ($log.Cells.Item($r,4).Value2.Length -gt 1000) 'Long message not on Log'}
 [void]$app.Run($macro,$book,'FBC',('x'*40000),$true)
 Check ($log.Range('D3').Value2.Length -eq 32767) 'Excel cell text limit was not handled'
 [void]$app.Run($macro,$book,'FBC','Latest complete',$false)
 Check ($log.Range('D3').Value2 -eq 'Latest complete' -and $log.Range('C3').Value2 -eq 'Complete') 'Latest message did not replace prior error'
 Check ($log.Cells.Item($log.Rows.Count,2).End(-4162).Row -eq 5) 'Log accumulated historical rows'
 $log.Activate();Check (-not $app.ActiveWindow.DisplayGridlines) 'Log gridlines should be hidden'
 $book.Save()
 Write-Output 'PASS native Log: all four operation slots, long messages, Excel cell limit, bounded latest-only output, unchanged Config values/G Stop Sign and fixed 30pt row heights.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
