Describe "Excel Worksheet" {
    Context "ExcelApplication" {
        It "Specified Worksheet" {
            $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
            $WorksheetName = 'DataTable3'
            $WorkSheetA = Get-ExcelWorksheetFromExcelApplication -ExcelPath $ExcelPath -WorksheetName $WorksheetName
            $RowCountA = $WorkSheetA.UsedRange.Rows.Count
            $ColummnCountA = $WorkSheetA.UsedRange.Columns.Count
            Close-ExcelWorksheet

            $WorkSheetB = Get-ExcelWorksheetFromImportExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName
            $RowCountB = $WorkSheetB.Dimension.Rows
            $ColummnCountB = $WorkSheetB.Dimension.Columns
            Close-ExcelWorksheet

            $RowCountA | Should -Be $RowCountB
            $ColummnCountA | Should -Be $ColummnCountB
        }

        It "Default Worksheet" {
            $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
            $WorkSheetA = Get-ExcelWorksheetFromExcelApplication -ExcelPath $ExcelPath
            $RowCountA = $WorkSheetA.UsedRange.Rows.Count
            $ColummnCountA = $WorkSheetA.UsedRange.Columns.Count
            Close-ExcelWorksheet

            $WorkSheetB = Get-ExcelWorksheetFromImportExcel -ExcelPath $ExcelPath
            $RowCountB = $WorkSheetB.Dimension.Rows
            $ColummnCountB = $WorkSheetB.Dimension.Columns
            Close-ExcelWorksheet

            $RowCountA | Should -Be $RowCountB
            $ColummnCountA | Should -Be $ColummnCountB
        }
    }
}