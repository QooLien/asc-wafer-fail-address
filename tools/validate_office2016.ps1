param(
    [Parameter(Mandatory = $true)][string]$SdkAssembly,
    [string]$Workbook
)
$ErrorActionPreference = 'Stop'
if (-not $Workbook) {
    $Workbook = Join-Path $PSScriptRoot '../templates/ASC_Fail_Address.xlsm'
}
Add-Type -AssemblyName WindowsBase
Add-Type -Path $SdkAssembly
$document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Workbook, $false)
try {
    if ($document.DocumentType -ne [DocumentFormat.OpenXml.SpreadsheetDocumentType]::MacroEnabledWorkbook) {
        throw ('Unexpected document type: ' + $document.DocumentType)
    }
    $validator = New-Object DocumentFormat.OpenXml.Validation.OpenXmlValidator([DocumentFormat.OpenXml.FileFormatVersions]::Office2016)
    $errors = @($validator.Validate([DocumentFormat.OpenXml.Packaging.OpenXmlPackage]$document))
    if ($errors.Count -gt 0) {
        $errors | ForEach-Object { Write-Output ($_.Description + ' | ' + $_.Part.Uri + ' | ' + $_.Path.XPath) }
        throw ('Office 2016 validation errors: ' + $errors.Count)
    }
    Write-Output 'PASS: Microsoft Open XML SDK Office2016 validation: 0 errors; MacroEnabledWorkbook.'
} finally {
    $document.Dispose()
}
