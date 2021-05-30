$StartPath = "$PSScriptRoot\.."
# JavaExcelBDD\src\test\resources\ExcelBDD.xlsx
$global:ExcelBDDFilePath = "$StartPath\..\JavaExcelBDD\src\test\resources\ExcelBDD.xlsx"
Get-Module MZExcel | Remove-Module
Import-Module $StartPath\MZExcel.psm1



Describe "Get BDD Data" {

    $BDDTestCaseList = Get-MZExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SimpleOpenBDD' `
        -ParameterNameColumn D `
        -HeaderRow 1

    It "Easy Success of Column List" -TestCases $BDDTestCaseList {
        Write-Host "Easy Success of Sheet $SheetName Column $Header"
        Write-Host "Header Row $HeaderRow"
        Write-Host "ParameterColumn $ParameterNameColumn"
        Write-Host "SheetName $SheetName"
        $IntHeaderRow = [int]$HeaderRow
        $TestcaseList = Get-MZExampleList -ExcelPath $ExcelBDDFilePath `
            -WorksheetName $SheetName `
            -ParameterNameColumn $ParameterNameColumn `
            -HeaderRow $IntHeaderRow `
            -HeaderMatcher $HeaderMatcher

        Write-Host ($TestcaseList | ConvertTo-Json )
        
        $TestcaseList[0]["Header"] | Should -Be $Header1Name
        $TestcaseList[0]["ParamName1"] | Should -Be $FirstGridValue
        $TestcaseList[3]["ParamName4"] | Should -Be $LastGridValue
        $TestcaseList[1]["ParamName1"] | Should -Be $ParamName1InSet2Value
        $TestcaseList[1]["ParamName2"] | Should -Be $ParamName2InSet2Value
        $TestcaseList[0]["ParamName3"] | Should -Be $ParamName3Value
        $TestcaseList[0].Count | Should -Be $ParameterCount
        $MaxBlankThreshold | Should -Be 3
        $TestcaseList.Count | Should -Be 4

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

