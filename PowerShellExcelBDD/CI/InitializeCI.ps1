Get-Module ExcelBDD | Remove-Module
Import-Module Pester
if (-Not (Get-InstalledModule -Name ImportExcel)) {
    Write-Host "Install ImportExcel"
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel
$global:StartPath = (Resolve-Path "$PSScriptRoot/..").Path
Write-Host $global:StartPath
$modulePath = Join-Path $StartPath "ExcelBDD/ExcelBDD.psm1"
Import-Module $modulePath