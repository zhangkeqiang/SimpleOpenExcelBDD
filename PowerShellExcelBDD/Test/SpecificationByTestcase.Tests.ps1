$global:StartPath = "$PSScriptRoot\.."
# JavaExcelBDD\src\test\resources\ExcelBDD.xlsx
$global:ExcelBDDFilePath = "$StartPath\..\JavaExcelBDD\src\test\resources\ExcelBDD.xlsx"
Get-Module MZExcel | Remove-Module
Import-Module $StartPath\MZExcel.psm1
Describe "Get Speicification by Example & Testcase" {
    
    $SpecificationByTestcaseList = Get-MZExampleWithTestResultList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SpecificationByTestcase' `
        -ParameterNameColumn E `
        -HeaderRow 4

    It "SpecificationByTestcase" -Testcases $SpecificationByTestcaseList {
        $Error.Clear()
        $Header | Should -Be $HeaderRowExpected
        $Error.Count | Should -Be 0
        ($HeaderRowTestResult -eq 'pass') | Should -Be $true

        $TestExcelPath = "$StartPath\..\JavaExcelBDD\src\test\resources\$ExcelFileName"
        $TestcaseList = Get-MZExampleWithTestResultList -ExcelPath $TestExcelPath `
            -WorksheetName  $SheetName `
            -ParameterNameColumn $ParameterNameColumn `
            -HeaderRow $HeaderRow
        
        [bool]$TestcaseList | Should -Be ($ExcelFileNameExpected -eq 'got')
        [bool]$TestcaseList | Should -Be ($SheetNameExpected -eq 'got')
        
    }
    
}