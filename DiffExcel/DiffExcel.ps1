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
if ($NewFile.IndexOf(":") -lt 0) {
    $NewFile = Join-Path $(Get-Location) $NewFile
}
if ($OldFile.IndexOf(":") -lt 0) {
    $OldFile = Join-Path $(Get-Location) $OldFile
}
Write-Output "File $NewFile"
Write-Output "Old File $OldFile"
Compare-Excel $OldFile $NewFile | Out-Null
# Show-Result $Result