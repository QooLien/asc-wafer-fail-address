param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
function RGB-Value($r,$g,$b){return $r+256*$g+65536*$b}
function Rebuild(){try{$msg=$app.Run("'"+$book.Name+"'!BLKVminEngine.BLK_BuildFailMaps",$book)}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw};$msg=''};Check ([string]::IsNullOrEmpty($msg)) $msg;$app.CalculateFull()}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false);$summary=$book.Worksheets.Item('BLK Fail Count');$before=$summary.Range('A1:I25').Value2
 Rebuild
 $after=$summary.Range('A1:I25').Value2
 for($r=1;$r -le 25;$r++){for($c=1;$c -le 9;$c++){Check ($before.GetValue($r,$c) -eq $after.GetValue($r,$c)) 'Map judgment/count changed'}}
 $map=$book.Worksheets.Item('Scan Fail Map')
 Check ($map.Range('C5').Value2 -eq 0 -and $map.Range('C5').Text -eq '') 'Zero must remain numeric0 but display blank'
 Check ($map.Range('C5').DisplayFormat.Interior.Color -eq (RGB-Value 255 255 255)) 'Zero must have white background'
 Check ($map.Range('C5').DisplayFormat.Font.Color -eq (RGB-Value 255 255 255)) 'Zero font must blend with background'
 Check ([string]::IsNullOrEmpty($map.Range('C6').Value2) -and $map.Range('C6').DisplayFormat.Interior.Color -eq (RGB-Value 191 191 191)) 'NoData gray was lost'
 Check ($map.Range('B5').Value2 -eq 3 -and $map.Range('F5').Value2 -eq 1) 'Expected counts3 and1'
 Check ($map.Range('B5').DisplayFormat.Interior.Color -eq (RGB-Value 248 105 107)) 'Highest count must be red'
 Check ($map.Range('F5').DisplayFormat.Interior.Color -eq (RGB-Value 255 235 132)) 'Count1 must be yellow'
 $scale=$null
 for($i=1;$i -le $map.Range('B5').FormatConditions.Count;$i++){if($map.Range('B5').FormatConditions.Item($i).Type -eq 3){$scale=$map.Range('B5').FormatConditions.Item($i)}}
 Check ($null -ne $scale -and $scale.ColorScaleCriteria.Count -eq 2) 'Must use native two-color scale'
 Check ($scale.ColorScaleCriteria.Item(1).Value -eq 1 -and $scale.ColorScaleCriteria.Item(2).Value -eq 3) 'Shared Pattern scale should use1..3'
 Check ($map.Range('A5').FormatConditions.Count -eq 0 -and $map.Range('B4').FormatConditions.Count -eq 0) 'Scale should not include axes'
 $map.Range('B5').Value2=[double]2;$app.CalculateFull();$middle=$map.Range('B5').DisplayFormat.Interior.Color
 Check ($middle -ne (RGB-Value 255 235 132) -and $middle -ne (RGB-Value 248 105 107)) 'Native scale did not interpolate automatically'
 $map.Range('B5').Value2=[double]3;$app.CalculateFull();$book.Save()
 # All-zero data must remain blank/white, with valid rules and gray actual missing records.
 $book.Worksheets.Item('FBC by BLK').Range('E2:IZ414').Value2=[double]0;Rebuild
 Check ($map.Range('B5').Value2 -eq 0 -and $map.Range('B5').Text -eq '' -and $map.Range('B5').DisplayFormat.Interior.Color -eq (RGB-Value 255 255 255)) 'All-zero case failed'
 Check ($map.Range('C6').DisplayFormat.Interior.Color -eq (RGB-Value 191 191 191)) 'All-zero case hid missing data'
 Write-Output 'PASS native Map heatmap: numeric0 blank/white font, positive1 yellow/3 red, true Excel2-color scale with live interpolation, same range across WT panels, axes excluded, missing gray retained, all-zero case valid and Map counts unchanged.'
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $app){$app.Quit()}}
