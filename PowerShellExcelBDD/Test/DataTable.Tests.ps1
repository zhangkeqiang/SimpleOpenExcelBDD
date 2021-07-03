
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
        $DataTableA[0]["Header01"] | Should -Be $FirstGridValue
        $DataTableA[5]["Header08"] | Should -Be $LastGridValue
        $DataTableA[2]["Header03"] | Should -Be $Header03InThirdSet

        $DataTableA[5].Count | Should -Be $ColumnCount
    }
}

Describe "Use ImportExcel Only" {
    It "Read From Excel" {
        $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
        $DataTable1 = Import-Excel -Path $ExcelPath -WorksheetName "DataTable1"  -StartRow 2 -StartColumn 1
        $DataTable1.Count | Should -Be 8
        Write-Host $DataTable1[0]."Header01"
        # $DataTableV05 = Import-Excel -Path $ExcelPath -WorksheetName 'DataTableV0.5' -NoHeader
    }
}


Describe "Use ExcelBDD to get DataTable" {
    $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTable1 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName DataTable1 -HeaderRow 2
    Show-ExampleList $DataTable1
    It "Use the DataTable" -Testcases $DataTable1 {

        Write-Host $Header01
        Write-Host $Header02
        Write-Host $Header08
        $Header01 | Should -Not -BeNullOrEmpty
    }
}