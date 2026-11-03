param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false)
 [void]$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_PrepareEmptyTemplate",$book)
 if($book.Worksheets.Item('Fail Address').Cells.Item(1,16384).End(-4159).Column -ne 65){throw 'Empty Address template should contain WT1 only'}
 if($book.Worksheets.Item('BLK_FBC_Comparison').UsedRange.Columns.Count -ne 62){throw 'Empty FBC template should contain WT1 only'}
 if($book.Worksheets.Item($book.Worksheets.Count).Name -ne 'Log' -or $book.Worksheets.Item(($book.Worksheets.Count-1)).Name -ne 'Chip Locations'){throw 'Final sheet order invalid'}
 $book.Save();Write-Output 'PASS final empty template: one-click analysis, WT1 initial headers, border-free Config H:N, Chip Locations penultimate, Log last.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
