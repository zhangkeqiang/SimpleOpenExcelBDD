
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
            # Write-Host $_
            $IsChanged = $true
            $Result[$Worksheet.Name] = "New worksheet"
        }
    }

    foreach ($Worksheet in $OldWorkBook.Worksheets) {
        try {
            if(-Not $NewWorkBook.Worksheets[$Worksheet.Name]){
                $IsChanged = $true
                $Result[$Worksheet.Name] = "Missing worksheet"
            }
        }
        catch {
            # Write-Host $_
            $IsChanged = $true
            $Result[$Worksheet.Name] = "Missing worksheet"
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
                        Cell = "$([char]($iColumn+64))$iRow"
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
                    Cell = "$([char]($iColumn+64))$iRow"
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
                        Cell = "$([char]($iColumn+64))$iRow"
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
                        Cell = "$([char]($iColumn+64))$iRow"
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
                Write-Host "Diff Cell:$($DiffItem.Cell), New:'$($DiffItem.New)', old:'$($DiffItem.Old)'"
            }
        }
    }
}