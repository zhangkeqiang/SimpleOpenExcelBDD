Describe "Specified Worksheet"  {
    It "Specified Worksheet From ExcelApplication " {
        $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
        $WorksheetName = 'DataTable3'
        $WorkSheetA = Get-ExcelWorksheetFromExcelApplication -ExcelPath $ExcelPath -WorksheetName $WorksheetName

        $WorkSheetA.Cells.Item(3,3).Text | Should -Be 'Header01'
        $WorkSheetA.Cells.Item(3,10).Text | Should -Be 'Header08'
        $WorkSheetA.Cells.Item(4,3).Text | Should -Be 'Value1.1'
        $WorkSheetA.Cells.Item(5,5).Text | Should -Be 'Value3.2'
        $WorkSheetA.Cells.Item(9,9).Text | Should -Be 'Value7.6'
        $WorkSheetA.Cells.Item(9,10).Text | Should -Be 'Value8.6'

        Close-ExcelWorksheet
    }

    It "Specified Worksheet From ExcelApplication " {
        $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
        $WorksheetName = 'DataTable3'
        $WorkSheetA = Get-ExcelWorksheetFromImportExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName

        $WorkSheetA.Cells.Item(3,3).Text | Should -Be 'Header01'
        $WorkSheetA.Cells.Item(3,10).Text | Should -Be 'Header08'
        $WorkSheetA.Cells.Item(4,3).Text | Should -Be 'Value1.1'
        $WorkSheetA.Cells.Item(5,5).Text | Should -Be 'Value3.2'
        $WorkSheetA.Cells.Item(9,9).Text | Should -Be 'Value7.6'
        $WorkSheetA.Cells.Item(9,10).Text | Should -Be 'Value8.6'

        Close-ExcelWorksheet
    }
}

Describe "Default Sheet"  -Tag 'DefaultSheet' {
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