
Describe "DataTable" {
    $ExcelBDDFilePath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTableList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'DataTableBDD' 

    It "Check DataTable Reading" -TestCases $DataTableList {
        $ExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
        $DataTableA = Get-DataTable -ExcelPath $ExcelPath -WorksheetName $SheetName `
            -HeaderRow $HeaderRow -StartColumn $StartColumn

        Show-ExampleList $DataTableA 
        $DataTableA.Count | Should -Be $TestSetCount
        $DataTableA.GetType().Name | Should -Be 'Object[]'
        $DataTableA[0].GetType().Name | Should -Be 'HashTable'
        $DataTableA[0]["Header1"] | Should -Be $FirstGridValue
        $DataTableA[5]["Header8"] | Should -Be $LastGridValue
        $DataTableA[5].Count | Should -Be $ColumnCount
    }
}

Describe "Use ImportExcel Only" {
    It "Read From Excel" {
        $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
        $DataTable1 = Import-Excel -Path $ExcelPath -WorksheetName "DataTable1"  -StartRow 2 -StartColumn 1
        $DataTable1.Count | Should -Be 6
        Write-Host $DataTable1[0]."Header1"
        # $DataTableV05 = Import-Excel -Path $ExcelPath -WorksheetName 'DataTableV0.5' -NoHeader
    }
}

Describe "Use ExcelBDD to get DataTable" {
    $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTable1 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName DataTable1 -HeaderRow 2
    Show-ExampleList $DataTable1
    It "Use the DataTable" -Testcases $DataTable1 {

        Write-Host $Header1
        Write-Host $Header2
        $Header1 | Should -Not -BeNullOrEmpty
    }
}