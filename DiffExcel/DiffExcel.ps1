[CmdletBinding()]
param (
    [Parameter()]
    [String]
    $OldFile,
    [Parameter()]
    [String]
    $NewFile
)
$StartPath = $PSScriptRoot
# Write-Host $StartPath
$modulePath = Join-Path $StartPath "DiffExcel.psm1"
Import-Module $modulePath
# Write-Output "Hello Diff Excel"
# Write-Output "Old File $OldFile"
$NewFile = Join-Path $(Get-Location) $NewFile
Write-Output "File $NewFile"
Compare-Excel $OldFile $NewFile | Out-Null
# Show-Result $Result