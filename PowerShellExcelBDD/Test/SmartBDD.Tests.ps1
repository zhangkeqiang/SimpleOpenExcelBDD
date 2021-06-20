$script:ExcelBDDFilePath = "$StartPath/BDDExcel/ExcelBDD.xlsx"

Describe "Get BDD Data only by sheet" {

    $BDDTestCaseList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SmartBDD' `
        -ParameterNameColumn F `
        -HeaderRow 1

    It "Get-SmartExampleList" -TestCases $BDDTestCaseList {
        Write-Host "Sheet $SheetName Column $Header"
        $TestcaseList = Get-SmartExampleList -ExcelPath $ExcelBDDFilePath `
            -WorksheetName $SheetName `
            -HeaderMatcher $HeaderMatcher `
            -HeaderUnmatcher $HeaderUnmatcher

        Write-Host ($TestcaseList | ConvertTo-Json )
        
        $TestcaseList[0]["Header"] | Should -Be $Header1Name
        $TestcaseList[0]["ParamName1"] | Should -Be $FirstGridValue
        $TestcaseList[3]["ParamName4"] | Should -Be $LastGridValue
        $TestcaseList[1]["ParamName1"] | Should -Be $ParamName1InSet2Value
        $TestcaseList[1]["ParamName2"] | Should -Be $ParamName2InSet2Value
        $TestcaseList[0]["ParamName3"] | Should -Be $ParamName3Value
        $TestcaseList[0].Count | Should -Be $ParameterCount
        $MaxBlankThreshold | Should -Be 3
        $TestcaseList.Count | Should -Be $TestDataSetCount

        $TestcaseList[0]["ParamName3"] | Should -Be ""
        $TestcaseList[1]["ParamName3"] | Should -Be ""
        $TestcaseList[2]["ParamName3"] | Should -Be ""
        $TestcaseList[3]["ParamName3"] | Should -Be ""

        $TestcaseList[0]["ParamName4"] | Should -Be "2021/4/30"
        $TestcaseList[1]["ParamName4"] | Should -Be "0"
        $TestcaseList[2]["ParamName4"] | Should -Be "1"
        $TestcaseList[3]["ParamName4"] | Should -Be "4.4"
    }
}