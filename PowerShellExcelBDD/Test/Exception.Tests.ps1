$script:ExcelBDDFilePath = "$StartPath/BDDExcel/ExcelBDD.xlsx"

Describe "Test Wrong Scenario" {

    $WrongFileList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'WrongFile' 

    It "Wrong File" -TestCases $WrongFileList {
        Write-Host "Header: $Header"
        $TestExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
        {
        $TestcaseList = Get-ExampleList -ExcelPath $TestExcelPath `
            -WorksheetName $SheetName `
            -HeaderMatcher $HeaderMatcher `
            -HeaderUnmatcher $HeaderUnmatcher
        } | Should -Throw  "*$SheetNameExpected*"
    }
}
