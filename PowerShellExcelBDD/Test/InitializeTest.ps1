$global:StartPath = (Resolve-Path "$PSScriptRoot/..").Path
Get-Module ExcelBDD | Remove-Module
$modulePath = Join-Path $StartPath "ExcelBDD.psd1"
Import-Module $modulePath