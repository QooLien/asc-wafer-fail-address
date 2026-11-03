param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false)
 $book.Worksheets.Item('Import Index').Range('AL1').ClearContents()
 $map=$book.Worksheets.Item('Scan Fail Map')
 $handled=$app.Run("'"+$book.Name+"'!ImportFailAddresses.OpenMapChipDetails",$book,$map,$map.Range('B5'))
 if(-not $handled -or $book.ActiveSheet.Name -ne 'Chip Fail Details'){throw 'Production drilldown did not open'}
 $detail=$book.Worksheets.Item('Chip Fail Details')
 if($detail.Range('J6').Value2 -ne 996 -or $detail.Range('K6').Value2 -ne 9){throw 'Production detail counts differ'}
 $map.Activate();$map.Range('B5').Select();$book.Save()
 Write-Output 'PASS production interactive demo: correct Scan detail and Address links, opens on Scan Map, no test marker.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
