param([string]$Workbook=(Join-Path $PSScriptRoot '../templates/ASC_ThreeASC_FBC.xlsm'),[string]$Output='C:/Users/dora9/outputs/Native_Address_Coverage_Test.xlsm')
# Compatibility entry point for the updated Map-only coverage behavior.
& (Join-Path $PSScriptRoot 'test_native_map_coverage.ps1') -Workbook $Workbook -Output $Output
