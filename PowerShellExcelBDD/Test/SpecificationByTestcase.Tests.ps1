# JavaExcelBDD\src\test\resources\ExcelBDD.xlsx
$script:ExcelBDDFilePath = "$StartPath/../JavaExcelBDD/src/test/resources/ExcelBDD.xlsx"

Describe "Get Speicification by Example & Testcase" {
    
    $SpecificationByTestcaseList = Get-TestcaseList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SpecificationByTestcase' `
        -ParameterNameColumn E `
        -HeaderRow 4

    It "SpecificationByTestcase" -Testcases $SpecificationByTestcaseList {
        $Error.Clear()

        $Header | Should -Be $HeaderRowExpected
        ($HeaderRowTestResult -eq 'pass') | Should -Be $true
        $HeaderRowTestResult | Should -Be 'pass'
        Write-Host "SheetName:"$SheetName
        Write-Host "ParameterNameColumn:"$ParameterNameColumn
        Write-Host "HeaderRow:"$HeaderRow

        $TestExcelPath = "$StartPath\..\JavaExcelBDD\src\test\resources\$ExcelFileName"
        $TestcaseList = Get-TestcaseList -ExcelPath $TestExcelPath `
            -WorksheetName  $SheetName `
            -ParameterNameColumn $ParameterNameColumn `
            -HeaderRow $HeaderRow
        
        $Error.Count | Should -Be 0

        [bool]$TestcaseList | Should -Be ($ExcelFileNameExpected -eq 'got')
        [bool]$TestcaseList | Should -Be ($SheetNameExpected -eq 'got')
        Write-Host ($TestcaseList | ConvertTo-Json -Depth 10)

        $TestcaseList[0]["$FirstSetFirstCheckedParam"] | Should -Be $FirstSetFirstCheckedParamExpected
        $TestcaseList[0]["ParamName1"] | Should -Be $FirstSetParamName1
        $TestcaseList[0]["ParamName2"] | Should -Be $FirstSetParamName2
        $TestcaseList[0]["ParamName3"] | Should -Be $FirstSetParamName3
        $TestcaseList[0]["ParamName4"] | Should -Be $FirstSetParamName4

        $TestcaseList[0]["ParamName1Expected"] | Should -Be $FirstSetParamName1Expected
        $TestcaseList[0]["ParamName2Expected"] | Should -Be $FirstSetParamName2Expected
        $TestcaseList[0]["ParamName3Expected"] | Should -Be $FirstSetParamName3Expected
        $TestcaseList[0]["ParamName4Expected"] | Should -Be $FirstSetParamName4Expected

        $TestcaseList[0]["ParamName1TestResult"] | Should -Be $FirstSetParamName1TestResult
        $TestcaseList[0]["ParamName2TestResult"] | Should -Be $FirstSetParamName2TestResult
        $TestcaseList[0]["ParamName3TestResult"] | Should -Be $FirstSetParamName3TestResult
        $TestcaseList[0]["ParamName4TestResult"] | Should -Be $FirstSetParamName4TestResult
    }
    
}