Get-Module DiffExcel | Remove-Module
$script:StartPath = (Resolve-Path "$PSScriptRoot\..").Path
Write-Host $StartPath
$modulePath = Join-Path $StartPath "DiffExcel.psm1"
Import-Module $modulePath

$ExcelBDDPath = Join-Path $StartPath "..\PowerShellExcelBDD\ExcelBDD\ExcelBDD.psm1"
Import-Module $ExcelBDDPath

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

Describe "Worksheet" {
    BeforeAll {
        $NewFile = "$PSScriptRoot\NewFile.xlsx"
        $OldFile = "$PSScriptRoot\OldFile.xlsx"
        $ExcelApp = New-Object -ComObject Excel.Application
        # $ExcelApp.Visible = $true
        $NewWorkBook = $ExcelApp.Workbooks.Open($NewFile)
        $OldWorkBook = $ExcelApp.Workbooks.Open($OldFile)
    }
    AfterAll {
        [void]$NewWorkBook.Close($false)
        [void]$OldWorkBook.Close($false)
        [void]$ExcelApp.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp)
    }

    $BDDExcelPath = "$PSScriptRoot\DiffExcelBDD.xlsx"
    $TestCaseList = Get-ExampleList -ExcelPath $BDDExcelPath `
        -WorksheetName DiffWorksheet -HeaderMatcher "Scenario3"
        
    It "Compare Worksheet" -TestCases $TestCaseList {
        Write-Host $WorksheetName
        Write-Host $DiffCount
        $Result = Compare-Worksheet $OldWorkBook.Worksheets[$WorksheetName] $NewWorkBook.Worksheets[$WorksheetName]
        $Result | ConvertTo-Json | Out-Host
        $Result.GetType().Name | Should -Be "Object[]"
        $Result.Count | Should -Be $DiffCount
        $Result[0] | Out-String | Should -Match $ResultText0
        $Result[1] | Out-String | Should -Match $ResultText1
        $Result[2] | Out-String | Should -Match $ResultText2
    }
}