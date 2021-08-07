Get-Module DiffExcel | Remove-Module
$StartPath = (Resolve-Path "$PSScriptRoot/..").Path
Write-Host $StartPath
$modulePath = Join-Path $StartPath "DiffExcel.psm1"
Import-Module $modulePath
Describe "DescribeName" {
    It "ItName" {
        $NewFile = "E:\Code\ExcelBDD\DiffExcel\Test\NewFile.xlsx"
        $OldFile = "E:\Code\ExcelBDD\DiffExcel\Test\OldFile.xlsx"
        $Result = Compare-Excel $OldFile $NewFile
        $Result | ConvertTo-Json -Depth 10 | Out-Host
        Show-Result $Result
    }
}