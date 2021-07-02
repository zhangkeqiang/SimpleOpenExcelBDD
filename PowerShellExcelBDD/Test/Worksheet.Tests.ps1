Describe "Excel Worksheet" {
    Context "ExcelApplication" {
        It "Get-ExcelWorksheetFromExcelApplication" {
            $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
            $WorksheetName = 'DataTable3'
            $WorkSheetA = Get-ExcelWorksheetFromExcelApplication -ExcelPath $ExcelPath -WorksheetName $WorksheetName
            $WorkSheetA.UsedRange.Rows.Count
            $WorkSheetA.UsedRange.Columns.Count
            Close-ExcelWorksheet

            $WorkSheetB = Get-ExcelWorksheetFromImportExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName
            $WorkSheetB.Dimension.Rows
            $WorkSheetB.Dimension.Columns
            Close-ExcelWorksheet
        }
    }
}