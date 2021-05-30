$StartPath = "$PSScriptRoot\.."
# JavaExcelBDD\src\test\resources\ExcelBDD.xlsx
$ExcelBDDFilePath = "$StartPath\..\JavaExcelBDD\src\test\resources\ExcelBDD.xlsx"
Get-Module MZExcel | Remove-Module
Import-Module $StartPath\MZExcel.psm1
Describe "Get Speicification by Example & Testcase" {
    
    $SpecificationByTestcaseList = Get-MZExampleWithTestResultList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SpecificationByTestcase' `
        -ParameterNameColumn E `
        -HeaderRow 4

    It "SpecificationByTestcase" -Testcases $SpecificationByTestcaseList {
        $Header | Should -Be ""
    }
    
}