
<#PSScriptInfo

.VERSION 1.1

.GUID c623da06-d9d6-4890-8171-627b0023c972

.AUTHOR zhangkq2000@hotmail.com

.COMPANYNAME ExcelBDD.com

.COPYRIGHT Copyright (c) 2021 by ExcelBDD Team, licensed under Apache 2.0 License.

.TAGS BDD ExcelBDD

.LICENSEURI https://www.apache.org/licenses/LICENSE-2.0.html

.PROJECTURI https://dev.azure.com/simplopen/ExcelBDD/_wiki/wikis/ExcelBDD.wiki/39/ExcelBDD-Homepage

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
now supports the following features:
Diff two excel files

#>

<# 
.DESCRIPTION 
 diff two excel files 
#> 
[CmdletBinding()]
param (
    [Parameter()]
    [String]
    $OldFile,
    [Parameter()]
    [String]
    $NewFile,
    [Switch]$Test
)
Write-Host "DiffExcel V1.0 developed by Zhang Keqiang for ExclBDD"
if ((-not $OldFile) -or (-not $NewFile)) {
    return
}
if ($NewFile.IndexOf(":") -lt 0) {
    $NewFile = Join-Path $(Get-Location) $NewFile
}
if ($OldFile.IndexOf(":") -lt 0) {
    $OldFile = Join-Path $(Get-Location) $OldFile
}
Write-Host "File $NewFile"
Write-Host "Old File $OldFile"
#Define functions

function Compare-Excel {
    param (
        $OldFile,
        $NewFile,
        [Switch]$Test
    )
    $ExcelApp = New-Object -ComObject Excel.Application
    # $ExcelApp.Visible = $true
    $NewWorkBook = $ExcelApp.Workbooks.Open($NewFile)
    $OldWorkBook = $ExcelApp.Workbooks.Open($OldFile)
    $Result = @{}
    $IsChanged = $false
    foreach ($Worksheet in $NewWorkBook.Worksheets) {
        # Write-Host $Worksheet.Name
        try {
            $Result[$Worksheet.Name] = Compare-Worksheet $OldWorkBook.Worksheets[$Worksheet.Name] $Worksheet
            if ($Result[$Worksheet.Name].GetType().Name -ne "String") {
                $IsChanged = $true
            }
        }
        catch {
            $IsChanged = $true
            # Write-Host $_
            $Result[$Worksheet.Name] = "New worksheet"
        }
    }
    if ($OldFile.IndexOf("Temp") -gt 0) {
        [void]$OldWorkBook.Close($false)
        $OpenWorkBook = $NewWorkBook
    }
    else {
        [void]$NewWorkBook.Close($false)
        $OpenWorkBook = $OldWorkBook
    }

    Show-Result $Result
    if ($IsChanged -and (-Not $Test)) {
        Start-Sleep 5
        $ExcelApp.Visible = $true
    }
    else {
        [void]$OpenWorkBook.Close($false)
        [void]$ExcelApp.Quit()
    }
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp)
    return $Result
}

function Compare-Worksheet {
    param (
        $OldWorksheet,
        $NewWorksheet
    )
    $DiffList = @()
    $NewRowsCount = $NewWorksheet.UsedRange.Row + $NewWorksheet.UsedRange.Rows.Count - 1
    # Write-Host "RowsCount $NewRowsCount"
    $NewColumnsCount = $NewWorksheet.UsedRange.Column + $NewWorksheet.UsedRange.Columns.Count - 1
    # Write-Host "ColumnsCount $NewColumnsCount"

    $OldRowsCount = $OldWorksheet.UsedRange.Row + $OldWorksheet.UsedRange.Rows.Count - 1
    # Write-Host "OldRowsCount $OldRowsCount"
    $OldColumnsCount = $OldWorksheet.UsedRange.Column + $OldWorksheet.UsedRange.Columns.Count - 1
    # Write-Host "OldColumnsCount $OldColumnsCount"
    
    for ($iRow = 1; $iRow -le $NewRowsCount; $iRow++) {
        for ($iColumn = 1; $iColumn -le $NewColumnsCount; $iColumn++) {
            try {
                if ($NewWorksheet.Cells.Item($iRow, $iColumn).Text -cne $OldWorksheet.Cells.Item($iRow, $iColumn).Text) {
                    $DiffItem = [PSCustomObject]@{
                        Grid = "$([char]($iColumn+64))$iRow"
                        # Row    = $iRow
                        # Column = $iColumn
                        Old  = $OldWorksheet.Cells.Item($iRow, $iColumn).Text
                        New  = $NewWorksheet.Cells.Item($iRow, $iColumn).Text
                    }
                    $DiffList += $DiffItem
                }
            }
            catch {
                # Write-Host $_
                $DiffItem = [PSCustomObject]@{
                    Grid = "$([char]($iColumn+64))$iRow"
                    # Row    = $iRow
                    # Column = $iColumn
                    Old  = $null
                    New  = $NewWorksheet.Cells.Item($iRow, $iColumn).Text
                }
                $DiffList += $DiffItem
            }
        }
    }

    if ($OldRowsCount -gt $NewRowsCount) {
        $MaxRowCount = $OldRowsCount
        for ($iRow = $NewRowsCount + 1; $iRow -le $OldRowsCount; $iRow++) {
            for ($iColumn = 1; $iColumn -le $NewColumnsCount; $iColumn++) {
                if (-Not [String]::IsNullOrWhiteSpace($OldWorksheet.Cells.Item($iRow, $iColumn).Text)) {
                    $DiffItem = [PSCustomObject]@{
                        Grid = "$([char]($iColumn+64))$iRow"
                        # Row    = $iRow
                        # Column = $iColumn
                        Old  = $OldWorksheet.Cells.Item($iRow, $iColumn).Text
                        New  = $null
                    }
                    $DiffList += $DiffItem
                }
            }
        }
    }
    else {
        $MaxRowCount = $NewRowsCount
    }

    if ($OldColumnsCount -gt $NewColumnsCount) {
        for ($iRow = 1; $iRow -le $MaxRowCount; $iRow++) {
            for ($iColumn = $NewColumnsCount + 1; $iColumn -le $OldColumnsCount; $iColumn++) {
                if (-Not [String]::IsNullOrWhiteSpace($OldWorksheet.Cells.Item($iRow, $iColumn).Text)) {
                    $DiffItem = [PSCustomObject]@{
                        Grid = "$([char]($iColumn+64))$iRow"
                        # Row    = $iRow
                        # Column = $iColumn
                        Old  = $OldWorksheet.Cells.Item($iRow, $iColumn).Text
                        New  = $null
                    }
                    $DiffList += $DiffItem
                }
            }
        }
    }
    if ($DiffList.Count -gt 0) {
        return $DiffList
    }
    return "No Change"
}

function Show-Result {
    param (
        $Result
    )
    foreach ($WorksheetName in $Result.Keys) {
        Write-Host "--- Worksheet $WorksheetName ---"
        foreach ($DiffItem in $Result[$WorksheetName]) {
            if ($DiffItem.GetType().Name -eq "String") {
                Write-Host $DiffItem
            }
            else {
                Write-Host "Diff Grid:$($DiffItem.Grid), New:'$($DiffItem.New)', old:'$($DiffItem.Old)'"
            }
        }
    }
}

#End of Define functions
Compare-Excel $OldFile $NewFile -Test:$Test | Out-Null
