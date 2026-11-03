param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false)
 try{$result=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_BuildComparison",$book)}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw};$result=''}
 if(-not [string]::IsNullOrEmpty($result)){throw $result}
 $ws=$book.Worksheets.Item('BLK_FBC_Comparison')
 if($ws.UsedRange.Rows.Count -ne 2 -or $ws.UsedRange.Columns.Count -ne 180){throw 'Expected empty template with180 columns and two header rows'}
 $book.Worksheets.Item('Config').Activate();$book.Save()
 Write-Output 'PASS native empty template: BLK_FBC_Comparison ready, three WT groups, each59 TNO columns; internal raw hidden.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
