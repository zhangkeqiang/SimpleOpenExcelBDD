Get-Module DiffExcel | Remove-Module
$script:StartPath = (Resolve-Path "$PSScriptRoot/..").Path
Write-Host $StartPath
$modulePath = Join-Path $StartPath "DiffExcel.psm1"
Import-Module $modulePath

Describe "Campare Whole File" {
    It "By Function" {
        $NewFile = "$PSScriptRoot\NewFile.xlsx"
        $OldFile = "$PSScriptRoot\OldFile.xlsx"
        $Result = Compare-Excel $OldFile $NewFile -Test
        # $Result | ConvertTo-Json -Depth 10 | Out-Host
        Write-Host "========================================================"
        Show-Result $Result
    }

    It "By Cmdlet" {
        $NewFile = "$PSScriptRoot\NewFile.xlsx"
        $OldFile = "$PSScriptRoot\OldFile.xlsx"
        & $StartPath\DiffExcel.ps1 $OldFile $NewFile -Test
    }
}

Describe "Compare Worksheet" {
    BeforeAll{
        $NewFile = "$PSScriptRoot\NewFile.xlsx"
        $OldFile = "$PSScriptRoot\OldFile.xlsx"
        $ExcelApp = New-Object -ComObject Excel.Application
        $ExcelApp.Visible = $true
        $NewWorkBook = $ExcelApp.Workbooks.Open($NewFile)
        $OldWorkBook = $ExcelApp.Workbooks.Open($OldFile)
    }
    AfterAll{
        [void]$NewWorkBook.Close($false)
        [void]$OldWorkBook.Close($false)
        [void]$ExcelApp.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp)
    }

    $BDDExcelPath = "$PSScriptRoot\DiffExcelBDD.xlsx"
    $TestCaseList = Get-ExampleList -ExcelPath $BDDExcelPath `
        -WorksheetName DiffWorksheet
        
    It "ItName" -TestCases $TestCaseList {
        Write-Host $WorksheetName
        Write-Host $ResultType
        Write-Host $DiffCount
        $Result = Compare-Worksheet $OldWorkBook.Worksheets[$WorksheetName] $NewWorkBook.Worksheets[$WorksheetName]
        $Result | ConvertTo-Json | Out-Host
        $Result.GetType().Name | Should -Be $ResultType
        $Result.Count | Should -Be $DiffCount
        $Result[$ResultNum1] | Out-String | Should -Match $ResultText1
        $Result[$ResultNum2] | Out-String | Should -Match $ResultText2
        $Result[$ResultNum3] | Out-String | Should -Match $ResultText3
    }
}