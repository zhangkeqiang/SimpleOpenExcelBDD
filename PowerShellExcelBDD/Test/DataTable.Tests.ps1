
Describe "DataTable" {
    $ExcelBDDFilePath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTableList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'DataTableBDD' `
        -ParameterNameColumn D `
        -HeaderRow 4 `

    It "Check DataTable Reading" -TestCases $DataTableList {
        $ExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
        $DataTable1 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName $SheetName -HeaderRow $HeaderRow
        $DataTable1.Count | Should -Be $TestSetCount
    }
}

Describe "ImportExcel" {
    It "Read From Excel" {
        $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
        $DataTable1 = Import-Excel -Path $ExcelPath -WorksheetName "DataTable1" -StartRow 2 -StartColumn 1
        $DataTable1.Count | Should -Be 6
        Write-Host $DataTable1[0]."Header1"
    }

    $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTable1 = Import-Excel -Path $ExcelPath -WorksheetName "DataTable1" -StartRow 2 -StartColumn 1

    It "Import" -Testcases $DataTable1 {
        Write-Host $Header1
    }
}