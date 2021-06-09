$global:StartPath = (Resolve-Path "$PSScriptRoot/..").Path
Get-Module ExcelBDD | Remove-Module
Import-Module $StartPath/ExcelBDD.psm1