param(
    [string]$Workbook = (Join-Path $PSScriptRoot '../templates/ASC_ThreeASC_FBC.xlsm'),
    [string]$Output = (Join-Path ([IO.Path]::GetTempPath()) 'Native_StopSign_Map_Test.xlsm'),
    [ValidateSet(64,128)][int]$ChipCount=128
)
$ErrorActionPreference = 'Stop'
Copy-Item -LiteralPath $Workbook -Destination $Output -Force
$app = $null; $book = $null
function Assert-Test($condition, [string]$message) { if (-not $condition) { throw $message } }
function RGB-Value([int]$r,[int]$g,[int]$b) { return $r + 256*$g + 65536*$b }
function Invoke-ExcelResult([string]$macro, $arg1, $arg2=$null) {
    try {
        if ($null -eq $arg2) { return $app.Run($macro,$arg1) }
        return $app.Run($macro,$arg1,$arg2)
    } catch [System.ArgumentNullException] {
        # Excel returns a null BSTR for a successful empty String result.
        # PowerShell7's COM marshaler rejects ptr=0; VBA callers accept it.
        if ($_.Exception.ParamName -eq 'ptr') { return '' }
        throw
    }
}
function FBC-Line([int]$tno,[int]$x,[int]$y,[string]$mode) {
    $values = for ($b=0;$b -lt 256;$b++) {
        if ($mode -eq 'all') { '1' }
        elseif ($mode -eq 'three' -and $b -eq 0) { '1' }
        elseif ($mode -eq 'three' -and $b -eq 17) { '3' }
        elseif ($mode -eq 'three' -and $b -eq 255) { '999999' }
        elseif ($mode -eq 'one' -and $b -eq 255) { '2' }
        else { '0' }
    }
    return "Tno= $tno DUT= 9 X= $x Y= $y FBC BY BLK= " + ($values -join ' ') + "`n"
}
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false; $app.DisplayAlerts = $false
    Write-Output ('Native Excel version=' + $app.Version)
    $book = $app.Workbooks.Open($Output,0,$false)
    $app.ScreenUpdating=$false; $app.EnableEvents=$false
    $cfg = $book.Worksheets.Item('Config')
    $cfg.Range('D2:G901').ClearContents()
    $cfg.Range('D2').Value2=[double]500; $cfg.Range('E2').Value2='Scan'; $cfg.Range('F2').Value2=[double]1.2
    $cfg.Range('D3').Value2=[double]501; $cfg.Range('E3').Value2='Scan'; $cfg.Range('F3').Value2=[double]1.035; $cfg.Range('G3').Value2='Scan'
    $cfg.Range('D4').Value2=[double]600; $cfg.Range('E4').Value2='SRAM'
    $cfg.Range('D5').Value2=[double]601; $cfg.Range('E5').Value2='SRAM'; $cfg.Range('F5').Value2=[double]0.765; $cfg.Range('G5').Value2='SRAM'
    $address = $book.Worksheets.Item('Fail Address')
    $address.Range('A2').Value2=[double]10; $address.Range('B2').Value2=[double]17; $address.Range('G2').Value2=[double]1
    $snapshot = $address.Range('A1:BT3').Value2
    $prefix = "'"+$book.Name+"'!BLKVminEngine."
    [void]$app.Run($prefix+'BLK_Setup',$cfg)
    $coords = @(Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../docs/coordinates.json') -Raw | ConvertFrom-Json | Select-Object -First $ChipCount)
    for ($phase=1;$phase -le 3;$phase++) {
        [void]$app.Run($prefix+'BLK_BeginFile',"synthetic-WT$phase")
        $raw = ''
        if ($phase -eq 1) {
            foreach ($xy in $coords) { $raw += 'DUT=1 X='+$xy[0]+' Y='+$xy[1]+" DC=Fail`n" }
        }
        $raw += FBC-Line 500 10 17 'all'
        if ($phase -eq 1) { $raw += FBC-Line 501 10 17 'zero' }
        if ($phase -eq 2) { $raw += FBC-Line 501 10 17 'three' }
        $raw += FBC-Line 501 12 17 'all'
        if ($phase -eq 1) { $raw += FBC-Line 601 12 17 'one' }
        if ($phase -eq 2) { $raw += FBC-Line 601 12 17 'zero' }
        [void]$app.Run($prefix+'BLK_ObserveLocations',$raw,$true)
        [void]$app.Run($prefix+'BLK_Feed',$raw,$true)
        $warning = Invoke-ExcelResult ($prefix+'BLK_WritePhase') $book ([int]$phase)
        Assert-Test ([string]::IsNullOrEmpty($warning)) ('FBC phase failed: '+$warning)
    }
    $warning=Invoke-ExcelResult ($prefix+'BLK_Finalize') $book
    Assert-Test ([string]::IsNullOrEmpty($warning)) ('Finalize failed: '+$warning)
    $warning=Invoke-ExcelResult ($prefix+'BLK_BuildFailMaps') $book
    Assert-Test ([string]::IsNullOrEmpty($warning)) ('Map build failed: '+$warning)
    $scan=$book.Worksheets.Item('Scan Fail Map')
    Assert-Test ($scan.Range('E12').Value2 -eq "(10,17)`n0") 'TNO500 incorrectly contaminated Stop501 Pass'
    Assert-Test ([int]$scan.Range('E12').Interior.Color -eq (RGB-Value 220 252 231)) 'Pass must be green'
    Assert-Test ($scan.Range('Q12').Value2 -eq "(10,17)`n3") 'WT2 Fail BLK count must be3'
    Assert-Test ([int]$scan.Range('Q12').Interior.Color -eq (RGB-Value 254 202 202)) 'Fail must be red'
    Assert-Test ($scan.Range('AC12').Value2 -eq "(10,17)`nNo Data") 'WT3 missing Stop record must remain No Data'
    Assert-Test ([int]$scan.Range('AC12').Interior.Color -eq (RGB-Value 191 191 191)) 'WT3 No Data must be gray'
    Assert-Test ($scan.Range('F12').Value2 -eq "(11,17)`nNo Data") 'DC-only Chip must be retained in Map'
    Assert-Test ($scan.Range('G12').Value2 -eq "(12,17)`n256") 'Full256 fail blocks missing'
    Assert-Test ($scan.Range('AE12').Value2 -eq "(12,17)`n256") 'Final complete FBC record at EOF was lost'
    $sram=$book.Worksheets.Item('SRAM Fail Map')
    Assert-Test ($sram.Range('G12').Value2 -eq "(12,17)`n1") 'WT1 final SRAM record was not flushed'
    Assert-Test ($sram.Range('S12').Value2 -eq "(12,17)`n0") 'WT2 final SRAM zero record was not flushed'
    $locations=$book.Worksheets.Item('Chip Locations')
    Assert-Test ($locations.Cells.Item($locations.Rows.Count,1).End(-4162).Row -eq ($ChipCount+1)) 'Raw coordinate inventory wrong'
    $summary=$book.Worksheets.Item('BLK Fail Count')
    Assert-Test ($summary.Cells.Item($summary.Rows.Count,1).End(-4162).Row -eq (2*$ChipCount*3+1)) '2 Patterns x ChipCount x3 WT summary wrong'
    $validOutput=Join-Path (Split-Path $Output -Parent) (([IO.Path]::GetFileNameWithoutExtension($Output))+'_Valid.xlsm')
    $book.SaveCopyAs($validOutput)
    Write-Output "PASS native normal Map:$ChipCount Chips,2 Patterns,3 WTs, counts0/3/256 and gray missing data."
    # A real parser error must be yellow Data Error, never gray untested.
    [void]$app.Run($prefix+'BLK_BeginFile','WT3-error')
    $bad="Tno=501 X=10 Y=17 FBC BY BLK=1 2 3`n"
    [void]$app.Run($prefix+'BLK_ObserveLocations',$bad,$true)
    [void]$app.Run($prefix+'BLK_Feed',$bad,$true)
    $warning=Invoke-ExcelResult ($prefix+'BLK_WritePhase') $book 3
    Write-Output ('Bad phase warning='+$warning)
    Assert-Test (-not [string]::IsNullOrEmpty($warning)) 'Selected malformed FBC was accepted'
    [void](Invoke-ExcelResult ($prefix+'BLK_Finalize') $book)
    $warning=Invoke-ExcelResult ($prefix+'BLK_BuildFailMaps') $book
    Assert-Test ([string]::IsNullOrEmpty($warning)) ('Rebuild with phase error failed: '+$warning)
    Write-Output ('WT3 status='+$locations.Range('E4').Value2+'; message='+$locations.Range('F4').Value2+'; MapCell='+$scan.Range('AC12').Value2)
    Assert-Test ($scan.Range('AC12').Value2 -eq "(10,17)`nData Error") 'Parse error must not look like No Data'
    Assert-Test ([int]$scan.Range('AC12').Interior.Color -eq (RGB-Value 255 217 102)) 'Data Error must be yellow'
    # Removing one Pattern prunes only its own previously generated Map.
    $cfg.Range('G5').ClearContents()
    $warning=Invoke-ExcelResult ($prefix+'BLK_BuildFailMaps') $book
    Assert-Test ([string]::IsNullOrEmpty($warning)) ('Pattern cleanup failed: '+$warning)
    $names=foreach($sheet in $book.Worksheets) { $sheet.Name }
    Assert-Test ($names -notcontains 'SRAM Fail Map') 'Obsolete managed Map was retained'
    $cfg.Range('G3').ClearContents()
    $warning=Invoke-ExcelResult ($prefix+'BLK_BuildFailMaps') $book
    Assert-Test ($warning -like 'Map skipped:*') 'No-Stop message missing'
    $names=foreach($sheet in $book.Worksheets) { $sheet.Name }
    Assert-Test ($names -notcontains 'Scan Fail Map') 'Removing all Stops left stale Map'
    Assert-Test ($summary.Cells.Item($summary.Rows.Count,1).End(-4162).Row -eq 1) 'Removing all Stops left stale counts'
    $after = $address.Range('A1:BT3').Value2
    for ($r=1;$r -le 3;$r++) { for ($c=1;$c -le 72;$c++) { Assert-Test ($snapshot.GetValue($r,$c) -eq $after.GetValue($r,$c)) "Address changed at $r,$c" } }
    $book.Save()
    Write-Output "PASS native Excel16: real FBC parse/output/sort,$ChipCount Raw XY, endpoint501-only counts0/3/256, three WT maps/colors, DC-only gray, parser-error yellow, managed Map cleanup and Address preservation."
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $app) { $app.Quit() }
    if ($null -ne $book) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
}
