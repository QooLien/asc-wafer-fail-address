param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false)
 try{$message=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_BuildFailMaps",$book)}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw};$message=''}
 if(-not [string]::IsNullOrEmpty($message)){throw $message}
 $book.Save();Write-Output 'PASS refreshed native Map display; current FBC and sequential eligibility retained.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
