
function Compare-Excel {
    param (
        $OldFile,
        $NewFile
    )
    $ExcelApp = New-Object -ComObject Excel.Application
    $ExcelApp.Visible = $true
    $NewWorkBook = $ExcelApp.Workbooks.Open($NewFile)
    $OldWorkBook = $ExcelApp.Workbooks.Open($OldFile)
    $Result = @{}
    foreach ($Worksheet in $NewWorkBook.Worksheets) {
        Write-Host $Worksheet.Name
        try {
            $Result[$Worksheet.Name] = Compare-Worksheet $OldWorkBook.Worksheets[$Worksheet.Name] $Worksheet
        }
        catch {
            Write-Host $_
            $Result[$Worksheet.Name] = "new worksheet"
            Write-Host "$($Worksheet.Name) is a new worksheet."
        }
    }

    [void]$NewWorkBook.Close($false)
    [void]$OldWorkBook.Close($false)
    [void]$ExcelApp.Quit() 
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp)
    return $Result
}

function Compare-Worksheet {
    param (
        $OldWorksheet,
        $NewWorksheet
    )
    $DiffList = @()
    $NewRowsCount = $NewWorksheet.UsedRange.Rows.Count
    Write-Host "RowsCount $NewRowsCount"
    $NewColumnsCount = $NewWorksheet.UsedRange.Columns.Count
    Write-Host "ColumnsCount $NewColumnsCount"

    $OldRowsCount = $OldWorksheet.UsedRange.Rows.Count
    Write-Host "OldRowsCount $OldRowsCount"
    $OldColumnsCount = $OldWorksheet.UsedRange.Columns.Count
    Write-Host "OldColumnsCount $OldColumnsCount"
    for ($iRow = 1; $iRow -le $NewRowsCount; $iRow++) {
        for ($iColumn = 1; $iColumn -le $NewColumnsCount; $iColumn++) {
            Write-Host $NewWorksheet.Cells.Item($iRow, $iColumn).Text
            try {
                if ($NewWorksheet.Cells.Item($iRow, $iColumn).Text -cne $OldWorksheet.Cells.Item($iRow, $iColumn).Text) {
                    $DiffItem = [PSCustomObject]@{
                        Grid   = "$([char]($iColumn+64))$iRow"
                        Row    = $iRow
                        Column = $iColumn
                        Old    = $OldWorksheet.Cells.Item($iRow, $iColumn).Text
                        New    = $NewWorksheet.Cells.Item($iRow, $iColumn).Text
                    }
                    $DiffList += $DiffItem
                }
            }
            catch {
                # Write-Host $_
                $DiffItem = [PSCustomObject]@{
                    Grid   = "$([char]($iColumn+64))$iRow"
                    Row    = $iRow
                    Column = $iColumn
                    Old    = $null
                    New    = $NewWorksheet.Cells.Item($iRow, $iColumn).Text
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
                        Grid   = "$([char]($iColumn+64))$iRow"
                        Row    = $iRow
                        Column = $iColumn
                        Old    = $OldWorksheet.Cells.Item($iRow, $iColumn).Text
                        New    = $null
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
                        Grid   = "$([char]($iColumn+64))$iRow"
                        Row    = $iRow
                        Column = $iColumn
                        Old    = $OldWorksheet.Cells.Item($iRow, $iColumn).Text
                        New    = $null
                    }
                    $DiffList += $DiffItem
                }
            }
        }
    }
    
    return $DiffList
}

function Show-Result {
    param (
        $Result
    )
    foreach ($WorksheetName in $Result.Keys) {
        Write-Host "--- Diff in $WorksheetName ---"
        $Result[$WorksheetName] | Format-Table | Out-Host
    }
}