$global:StartPath = (Resolve-Path "$PSScriptRoot/..").Path
Get-Module ExcelBDD | Remove-Module
$modulePath = Join-Path $StartPath "Module/ExcelBDD.psd1"
Import-Module $modulePath