$script:ExcelBDDFilePath = "$StartPath/BDDExcel/ExcelBDD.xlsx"

Describe "Get-ExampleListByHeader" {

    $BDDTestCaseList = Get-ExampleListByHeader -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SpecificationByExample' `
        -ParameterNameColumn F `
        -HeaderRow 1 `
        -HeaderMatcher Scenario

    It "Easy Success of Column List" -TestCases $BDDTestCaseList {
        Write-Host "Header $Header"
        Write-Host "SheetName $SheetName"
        Write-Host "HeaderRow $HeaderRow"
        Write-Host "ParameterColumn $ParameterNameColumn"
        Write-Host "HeaderMatcher $HeaderMatcher"
        Write-Host "HeaderUnmatcher $HeaderUnmatcher"
        Write-Host "Expected:($ExpectedSwitch -eq 'On')"
        Write-Host "TestResult:($TestResultSwitch -eq 'On')"

        $TestcaseList = Get-ExampleListByHeader -ExcelPath $ExcelBDDFilePath `
            -WorksheetName $SheetName `
            -ParameterNameColumn $ParameterNameColumn `
            -HeaderRow $HeaderRow `
            -HeaderMatcher $HeaderMatcher `
            -HeaderUnmatcher $HeaderUnmatcher `
            -Expected:($ExpectedSwitch -eq 'On') `
            -TestResult:($TestResultSwitch -eq 'On')

        Write-Host ($TestcaseList | ConvertTo-Json )
        
        $TestcaseList[0]["Header"] | Should -Be $Header1Name
        $TestcaseList[0]["ParamName1"] | Should -Be $FirstGridValue
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
        $TestcaseList[3]["ParamName4"] | Should -Be $LastGridValue
    }
}