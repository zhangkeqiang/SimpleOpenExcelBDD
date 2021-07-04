
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
        $DataTableA[5].Count | Should -Be $ColumnCount
        # one check is added for V0.5
        $DataTableA[2]["Header03"] | Should -Be $Header03InThirdSet
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


Describe "Use ExcelBDD to get DataTable V0.4" {
    #get hashtable list as a data table
    $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTable1 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName DataTable1 -HeaderRow 2 -StartColumn 'A'
    Show-ExampleList $DataTable1
    #Then use the hashtable list
    It "Use the DataTable V0.4" -Testcases $DataTable1 {
        $Header01 | Should -Match "^Value1"
        $Header02 | Should -Match "^Value2"
        $Header03 | Should -Match "^Value3"
        $Header04 | Should -Match "^Value4"
        $Header05 | Should -Match "^Value5"
        $Header06 | Should -Match "^Value6"
        $Header07 | Should -Match "^Value7"
        $Header08 | Should -Match "^Value8"
    }
}

Describe "Use ExcelBDD to get DataTable V0.5" {
    $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    #get hashtable list as a data table, if StartColumn is 1, it can be ignored
    $DataTableV05 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName "DataTableV0.5" -HeaderRow 2
    Show-ExampleList $DataTableV05
    #Then use the hashtable list
    It "Use the DataTable V0.5" -Testcases $DataTableV05 {
        $Header01 | Should -Match "^Value1"
        $Header02 | Should -Match "^Value2"
        $Header03 | Should -Match "^Value3"
        $Header04 | Should -Match "^Value4"
        $Header05 | Should -Match "^Value5"
        $Header06 | Should -Match "^Value6"
        $Header07 | Should -Match "^Value7"
        $Header08 | Should -Match "^Value8"
    }
}