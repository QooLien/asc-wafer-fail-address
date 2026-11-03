param(
    [string]$Workbook=(Join-Path $PSScriptRoot '../templates/ASC_ThreeASC_FBC.xlsm'),
    [string]$Output=(Join-Path ([IO.Path]::GetTempPath()) 'Native_Function_BLK_Audit_Test.xlsm')
)
$ErrorActionPreference='Stop'
Copy-Item -LiteralPath $Workbook -Destination $Output -Force
$app=$null; $book=$null; $sourceBook=$null
function Assert-Audit($condition,[string]$message) { if(-not $condition) { throw $message } }
function Equal-Grid($a,$b,[int]$rows,[int]$cols) {
    for($r=1;$r -le $rows;$r++){for($c=1;$c -le $cols;$c++){if($a.GetValue($r,$c) -ne $b.GetValue($r,$c)){throw "Source changed at $r,$c"}}}
}
try {
    $app=New-Object -ComObject Excel.Application
    $app.Visible=$false; $app.DisplayAlerts=$false
    $book=$app.Workbooks.Open($Output,0,$false)
    $app.ScreenUpdating=$false; $app.EnableEvents=$false
    Write-Output ('Native audit Excel version='+$app.Version)
    $cfg=$book.Worksheets.Item('Config');$address=$book.Worksheets.Item('Fail Address');$fbc=$book.Worksheets.Item('FBC by BLK');$index=$book.Worksheets.Item('Import Index')
    $cfg.Range('D2:G901').ClearContents()
    $cfg.Range('D2').Value2=[double]500;$cfg.Range('E2').Value2='Scan500';$cfg.Range('F2').Value2=[double]1.2
    $cfg.Range('D3').Value2=[double]501;$cfg.Range('E3').Value2='Scan501';$cfg.Range('F3').Value2=[double]1.2;$cfg.Range('G3').Value2='Scan'
    $cfg.Range('Q2').Value2='All Function'
    $address.UsedRange.ClearContents()
    $head=New-Object 'object[,]' 1,12
    $base=@('X','Y','BLK','WL','BL','IO')
    for($c=0;$c -lt 6;$c++){$head[0,$c]=$base[$c]}
    for($phase=1;$phase -le 3;$phase++){for($i=0;$i -lt 2;$i++){$head[0,(6+($phase-1)*2+$i)]="WT$phase`nScan$(500+$i)_1.20v"}}
    $address.Range('A1:L1').Value2=$head
    $detail=New-Object 'object[,]' 1002,12
    for($r=0;$r -lt 1000;$r++){
        $detail[$r,0]=[double]10;$detail[$r,1]=[double]17;$detail[$r,2]=[double]0;$detail[$r,3]=[double]$r;$detail[$r,4]=[double]0;$detail[$r,5]=[double]0;$detail[$r,7]=[double]1
    }
    $detail[1000,0]=[double]10;$detail[1000,1]=[double]17;$detail[1000,2]=[double]1;$detail[1000,3]=[double]0;$detail[1000,4]=[double]0;$detail[1000,5]=[double]0;$detail[1000,6]=[double]1
    $detail[1001,0]=[double]10;$detail[1001,1]=[double]17;$detail[1001,2]=[double]2;$detail[1001,3]=[double]0;$detail[1001,4]=[double]0;$detail[1001,5]=[double]0;$detail[1001,9]=[double]1
    $address.Range('A2:L1003').Value2=$detail
    $fbc.UsedRange.ClearContents()
    $fh=New-Object 'object[,]' 1,260
    for($c=0;$c -lt 4;$c++){$fh[0,$c]=@('WT','TNO','X','Y')[$c]}
    for($b=0;$b -lt 256;$b++){$fh[0,($b+4)]="BLK $b"}
    $fbc.Range('A1:IZ1').Value2=$fh
    $records=New-Object 'object[,]' 5,260
    $wts=@('WT1','WT1','WT2','WT3','WT1');$tnos=@(501,500,501,501,350)
    for($r=0;$r -lt 5;$r++){
        $records[$r,0]=$wts[$r];$records[$r,1]=[double]$tnos[$r];$records[$r,2]=[double]10;$records[$r,3]=[double]17
        for($b=4;$b -lt 260;$b++){$records[$r,$b]=[double]0}
    }
    $records[0,4]=[double]4;$records[0,5]=[double]6;$records[0,259]=[double]2147483648
    $records[1,5]=[double]3;$records[2,6]=[double]1;$records[3,6]=[double]1;$records[4,7]=[double]99
    $fbc.Range('A2:IZ6').Value2=$records
    for($phase=0;$phase -lt 3;$phase++){
        $index.Cells.Item((2+$phase),8).Value2='Address';$index.Cells.Item((5+$phase),8).Value2='FBC'
        foreach($row in @((2+$phase),(5+$phase))){
            $index.Cells.Item($row,9).Value2="WT$($phase+1)";$index.Cells.Item($row,10).Value2='TESTLOT';$index.Cells.Item($row,11).Value2='07';$index.Cells.Item($row,12).Value2="WTT$($phase+1)";$index.Cells.Item($row,13).Value2='TEST_VER_FA'
        }
        $book.Worksheets.Item('Chip Locations').Cells.Item((2+$phase),5).Value2='OK'
    }
    $addrBefore=$address.Range('A1:L1003').Value2;$fbcBefore=$fbc.Range('A1:IZ6').Value2
    $macro="'"+$book.Name+"'!ImportFailAddresses.AuditFunctionBlocks"
    $result=$app.Run($macro,$book)
    Assert-Audit ($result -like '*rows: 3') ('Unexpected audit result: '+$result)
    $report=$book.Worksheets.Item('Missing BLK Address')
    $found=@{}
    for($r=5;$r -le 7;$r++){
        $key=[string]$report.Cells.Item($r,1).Value2+'|'+$report.Cells.Item($r,2).Value2+'|'+$report.Cells.Item($r,8).Value2
        $found[$key]=$true
        Assert-Audit ($report.Cells.Item($r,2).Value2 -eq 501) 'Monitor or another TNO was reported'
        if($report.Cells.Item($r,1).Value2 -eq 'WT1'){
            Assert-Audit ($report.Cells.Item($r,10).Value2 -eq 1000) 'Recorded cap count incorrect'
            Assert-Audit ($report.Cells.Item($r,11).Value2 -eq 'Yes') '>=1000 hint missing'
        }
    }
    Assert-Audit ($found.ContainsKey('WT1|501|1') -and $found.ContainsKey('WT1|501|255') -and $found.ContainsKey('WT3|501|2')) 'Exact WT/TNO/Chip/BLK keys wrong'
    Assert-Audit ($report.Cells.Item(6,9).Value2 -eq 2147483648) 'Large FBC overflowed Long'
    Equal-Grid $addrBefore ($address.Range('A1:L1003').Value2) 1003 12
    Equal-Grid $fbcBefore ($fbc.Range('A1:IZ6').Value2) 6 260
    # Existing worksheet filters must not hide underlying Address evidence.
    [void]$address.Range('A1:L1003').AutoFilter(3,0)
    $result=$app.Run($macro,$book);Write-Output ('Filtered source result='+$result);Assert-Audit ($result -like '*rows: 3') 'Visible-row filtering affected comparison'
    $address.AutoFilterMode=$false
    [void]$fbc.Range('A1:IZ6').AutoFilter(2,350)
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*rows: 3') 'FBC filter hid Function evidence'
    $fbc.AutoFilterMode=$false
    $cfg.Range('Q2').Value2='Stop Sign';$cfg.Range('G2').Value2='Only500';$cfg.Range('G3').ClearContents()
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*rows: 0') ('No-gap case failed: '+$result)
    Assert-Audit ($book.Worksheets.Item('Missing BLK Address').Cells.Item(1048576,1).End(-4162).Row -eq 4) 'No-gap report retained old rows'
    $cfg.Range('G2').ClearContents();$cfg.Range('G3').Value2='Scan'
    # One Address row + one requested column per WT exercises scalar Value2.
    $cfg.Range('Q2').Value2='Stop Sign';$address.Range('A3:L1003').ClearContents()
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*rows: 4') ('Single-row/Stop scope failed: '+$result)
    $address.Range('A1:L1003').Value2=$addrBefore;$cfg.Range('Q2').Value2='All Function'
    # Duplicate human labels require an exact numeric snapshot, never guessed positions.
    $cfg.Range('E2:E3').Value2='Scan'
    for($phase=1;$phase -le 3;$phase++){for($i=0;$i -lt 2;$i++){$address.Cells.Item(1,(7+($phase-1)*2+$i)).Value2="WT$phase`nScan_1.20v"}}
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*ambiguous*') 'Ambiguous legacy labels were accepted'
    $mapping=New-Object 'object[,]' 6,4
    for($phase=1;$phase -le 3;$phase++){for($i=0;$i -lt 2;$i++){
        $r=($phase-1)*2+$i;$c=7+$r;$mapping[$r,0]="WT$phase";$mapping[$r,1]=[string](500+$i);$mapping[$r,2]=[double]$c;$mapping[$r,3]="WT$phase`nScan_1.20v"
    }}
    $index.Range('A2:D7').Value2=$mapping
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*rows: 3') ('Numeric TNO snapshot failed: '+$result)
    $report=$book.Worksheets.Item('Missing BLK Address');$savedTitle=$report.Range('A1').Value2
    $index.Range('J5').Value2='OTHERLOT'
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*mismatch*') 'Mismatched sources accepted'
    Assert-Audit ($book.Worksheets.Item('Missing BLK Address').Range('A1').Value2 -eq $savedTitle) 'Failed audit replaced previous report'
    $index.Range('J5').Value2='TESTLOT';$index.Range('H2').Value2='Address Pending'
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*incomplete*') 'Incomplete import accepted'
    $index.Range('H2').Value2='Address'
    $book.Worksheets.Item('Chip Locations').Range('E3').Value2='Data Error'
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*parse/output error*') 'FBC parse error accepted as cap gap'
    $book.Worksheets.Item('Chip Locations').Range('E3').Value2='OK'
    $fbc.Range('A7:IZ7').Value2=$fbc.Range('A2:IZ2').Value2
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*Duplicate FBC*') 'Duplicate FBC rows accepted'
    Assert-Audit ($book.Worksheets.Count -eq 7) 'Failed staging worksheet leaked'
    $fbc.Range('A7:IZ7').ClearContents()
    $result=$app.Run($macro,$book);Assert-Audit ($result -like '*rows: 3') 'Final audit failed'
    # Read-only legacy/result file can be audited into a separate tool workbook.
    $sourcePath=Join-Path (Split-Path $Output -Parent) 'Native_Audit_Source_Result.xlsm'
    $book.SaveCopyAs($sourcePath)
    $sourceBook=$app.Workbooks.Open($sourcePath,0,$true)
    $result=$app.Run($macro,$sourceBook,$book)
    Assert-Audit ($result -like '*rows: 3') ('Read-only source audit failed: '+$result)
    Assert-Audit ($sourceBook.ReadOnly) 'Source was not opened read-only'
    $sourceBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sourceBook);$sourceBook=$null
    Equal-Grid $fbcBefore ($fbc.Range('A1:IZ6').Value2) 6 260
    Equal-Grid $addrBefore ($address.Range('A1:L1003').Value2) 1003 6
    $finalAddress=$address.Range('A1:L1003').Value2
    for($r=2;$r -le 1003;$r++){for($c=7;$c -le 12;$c++){Assert-Audit ($addrBefore.GetValue($r,$c) -eq $finalAddress.GetValue($r,$c)) "Address flag changed at $r,$c"}}
    $book.Save()
    Write-Output 'PASS native Function BLK audit: exact WT/TNO/Chip/BLK,1000-address cap hint, BLK255/large FBC, Monitor exclusion, hidden rows, Stop scope, scalar row, numeric/legacy header safety, source/pending/error checks, transactional report and read-only old-result input; Address/FBC unchanged.'
} finally {
    if($null -ne $sourceBook){$sourceBook.Close($false)}
    if($null -ne $book){$book.Close($false)}
    if($null -ne $app){$app.Quit()}
    if($null -ne $book){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($null -ne $app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
