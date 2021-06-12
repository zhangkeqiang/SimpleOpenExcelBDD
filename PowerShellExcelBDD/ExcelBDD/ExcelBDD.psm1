#Define max blank lines for ending sheet reading
$MaxBlankThreshold = 3

<#
.Description
Get worksheet from Excel file according to build sheet path and worksheet name
.Example
Get-MZExcelWorksheet -ExcelPath C:\buildsheet.xlsx -WorksheetName 'PaaS SQL DB Build'
#>
function Get-MZExcelWorksheet {
    param (
        [String]$ExcelPath,
        [String]$WorksheetName
    )
    if (-not (Test-Path $ExcelPath)) {
        throw "$ExcelPath file doesn't exist."
    }
    try {
        $script:appExcel = Open-ExcelPackage -Path $ExcelPath
        if ($WorksheetName) {
            $Worksheet = $appExcel.Workbook.Worksheets[$WorksheetName]
        }
        else {
            $Worksheet = $appExcel.Workbook.Worksheets | Select-Object -First 1
        }
    }
    catch {
        $script:appExcel = New-Object -ComObject Excel.Application
        # Let Excel run in the backend, comment out below line, if debug, remove below #
        # $script:appExcel.Visible = $true
        $WorkBook = $script:appExcel.Workbooks.Open($ExcelPath)
        $Worksheet = $WorkBook.Sheets[$WorksheetName]
    }
    return $Worksheet
}

function Close-MZExcelWorksheet {
    try {
        if ($script:appExcel.Name -eq "Microsoft Excel") {
            $script:appExcel.ActiveWorkbook.Close($false)
            $script:appExcel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:appExcel)
        }
        else {
            Close-ExcelPackage -ExcelPackage $script:appExcel -NoSave
        }
    }
    catch {
        Write-Debug "Excel is closed."
    }
}
<#
.Description
Get worksheet information as a hashtable list according to Header Mapping
#>
# function Get-MZHashTableListFromWorksheet {
#     param (
#         $Worksheet,
#         $HeaderMapping,
#         $MandatoryColumnNum = 1,
#         $StartRow = 3,
#         $MaxRow = 100
#     )
#     $List = [System.Collections.ArrayList]::new()
#     for ($iRow = $StartRow; $iRow -lt $MaxRow; $iRow++) {
#         if (Test-MZHasValue $Worksheet.Cells.Item($iRow, $MandatoryColumnNum).Text) {
#             #This Row has values
#             $RowSet = @{}
#             for ($iCol = 1; $iCol -lt $HeaderMapping.count; $iCol++) {
#                 if (Test-MZHasValue $HeaderMapping[$iCol][1]) {
#                     $RowSet[$HeaderMapping[$iCol][1]] = $Worksheet.Cells.Item($iRow, $iCol).Text
#                 }
#             }
#             [void]$List.Add($RowSet)
#         }
#     }
#     Close-MZExcelWorksheet | Out-Null
#     return $List
# }



<#
.Description
Get a Hashtable list from excel sheet, one row for one hashtable
.Example
    #Get TestcaseDataList for Pester Testcase
    $TestcaseDataList = Get-MZHashTableListFromExcel -WorksheetName SheetName `
        -ExcelPath "${StartPath}${SEP}IaCSQLDBToolKit${SEP}TestData${SEP}DBTestCaseData.xlsx"  `
        -HeaderRow 1
    It "Full Rule Except Email From Excel File" -Testcases $TestcaseDataList {
        Test-MZIsPropertyValid -PropertyName $PropertyName -PropertyValue $PropertyValue -Rule $Rule | Should -Be ($Expected -eq "TRUE")
    }
#>
# function Get-MZHashTableListFromExcel {
#     param (
#         [String]$ExcelPath,
#         [String]$WorksheetName,
#         $MandatoryColumnNum = 1,
#         $HeaderRow = 1
#     )
#     $MaxRow = 1000
#     $MaxCol = 100
#     $Worksheet = Get-MZExcelWorksheet -ExcelPath $ExcelPath -WorksheetName $WorksheetName
#     if ($null -eq $Worksheet ) {
#         Write-MZDebug "'$WorksheetName' sheet doesn't exist in $ExcelPath."
#         return $null
#     }
#     Write-MZDebug "'$WorksheetName' sheet exists in $ExcelPath."
#     $List = [System.Collections.ArrayList]::new()
#     $StartRow = $HeaderRow + 1
#     for ($iRow = $StartRow; $iRow -lt $MaxRow; $iRow++) {
#         if (Test-MZHasValue $Worksheet.Cells.Item($iRow, $MandatoryColumnNum).Text) {
#             #This Row has values
#             $RowSet = @{}
#             for ($iCol = 1; $iCol -lt $MaxCol; $iCol++) {
#                 if (Test-MZHasValue $Worksheet.Cells.Item($HeaderRow, $iCol).Text) {
#                     $RowSet[$Worksheet.Cells.Item($HeaderRow, $iCol).Text.Trim()] = $Worksheet.Cells.Item($iRow, $iCol).Text
#                 }
#                 else {
#                     break
#                 }
#             }
#             [void]$List.Add($RowSet)
#         }
#         else {
#             break
#         }
#     }
#     Close-MZExcelWorksheet | Out-Null
#     return $List
# }

<#
.Description
Get Specification By Example Data from Excel ,input and output are separated in columns
#>
# function Get-MZExampleList2 {
#     param (
#         [String]$ExcelPath,
#         [String]$WorksheetName,
#         $MandatoryRowNum = 1,
#         $ParamNameCol = 3,
#         $StartRow = 2,
#         $MaxRow = 1000,
#         $MaxCol = 100
#     )
#     $Worksheet = Get-MZExcelWorksheet -ExcelPath $ExcelPath -WorksheetName $WorksheetName
#     if ($null -eq $Worksheet ) {
#         return $null
#     }
#     $StartCol = $ParamNameCol + 1
#     $List = [System.Collections.ArrayList]::new()
#     for ($iCol = $StartCol; $iCol -lt $MaxCol; $iCol += 2) {
#         if ([String]::IsNullOrEmpty($Worksheet.Cells.Item($MandatoryRowNum, $iCol).Text)) {
#             #This Row has no value
#             break
#         }
#         else {
#             $DataSet = [ordered]@{}
#             for ($iRow = $StartRow; $iRow -lt $MaxRow; $iRow++) {
#                 if ([String]::IsNullOrEmpty($Worksheet.Cells.Item($iRow, $ParamNameCol).Text)) {
#                     break
#                 }
#                 else {
#                     $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim()] = $Worksheet.Cells.Item($iRow, $iCol).Text
#                     $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim() + "Expected"] = $Worksheet.Cells.Item($iRow, $iCol + 1).Text
#                 }
#             }
#             [void]$List.Add($DataSet)
#         }
#     }
#     Close-MZExcelWorksheet | Out-Null
#     return $List
# }

<#
.Description
Get BDD/Specification By Example Data from Excel ,input, output and test result are separated in 3 columns
#>
function Get-TestcaseList {
    param (
        [String]$ExcelPath,
        [String]$WorksheetName,
        $HeaderRow = 1,
        $ParameterNameColumn = 'C'
    )
    $MaxRow = 1000
    $MaxCol = 100
    $Worksheet = Get-MZExcelWorksheet -ExcelPath $ExcelPath -WorksheetName $WorksheetName
    if ($null -eq $Worksheet ) {
        return $null
    }
    if ($HeaderRow.GetType().Name -eq 'String') {
        $HeaderRow = [int]$HeaderRow
    }
    $StartRow = $HeaderRow + 1
    $ParamNameCol = [int][char]($ParameterNameColumn.ToUpper()) - 64
    $StartCol = $ParamNameCol + 1

    $List = [System.Collections.ArrayList]::new()
    for ($iCol = $StartCol; $iCol -lt $MaxCol; $iCol += 3) {
        if ([String]::IsNullOrEmpty($Worksheet.Cells.Item($HeaderRow, $iCol).Text)) {
            #This Row has no value
            break
        }
        else {
            $DataSet = [ordered]@{}
            $DataSet["Header"] = $Worksheet.Cells.Item($HeaderRow - 1, $iCol).Text
            for ($iRow = $StartRow; $iRow -lt $MaxRow; $iRow++) {
                if ([String]::IsNullOrEmpty($Worksheet.Cells.Item($iRow, $ParamNameCol).Text)) {
                    break
                }
                else {
                    $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim()] = $Worksheet.Cells.Item($iRow, $iCol).Text
                    $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim() + "Expected"] = $Worksheet.Cells.Item($iRow, $iCol + 1).Text
                    $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim() + "TestResult"] = $Worksheet.Cells.Item($iRow, $iCol + 2).Text
                }
            }
            [void]$List.Add($DataSet)
        }
    }
    Close-MZExcelWorksheet | Out-Null
    return $List
}

<#
.Description
Get hashtable list of Example data, one Hashtable from one column in excel sheet

.Example
    use default HeaderRow which is 1, and default ParameterNameColumn which is C
    $ExampleList = Get-ExampleList -ExcelPath ".\Excel\Example1.xlsx" -WorksheetName 'Scenario1'
    It "Easy Success of SBE" -TestCases $ExampleList {
        [int]$BlackSweaterCountAtCustomer | Should -BeGreaterOrEqual $BlackSweaterCountReturned
        [int]$BlackSweaterCountInInvertory1 + [int]$BlackSweaterCountReturned | Should -Be $BlackSweaterCountInInvertory2
    }

    Describe "Test filter the dashboard by department" {
    $ExampleList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'StoryExample1' `
        -ParameterNameColumn E `
        -HeaderRow 3
    It "Run Example one by one" -TestCases $ExampleList {
        #The below variables are generated automatically from Excel
        Write-Host "===$Header==="
        Write-Host $SelectedView
        Write-Host $DepartmentCount
        Write-Host $SelectedDepartment
        Write-Host $FullDepartmentName
        Write-Host $DepartmentLocation
        Write-Host $DepartmentCurrentMonthKPI1
        Write-Host $DepartmentCurrentMonthKPI2
    }
}
    


#>
function Get-ExampleList {
    param (
        [String]$ExcelPath,
        [String]$WorksheetName,
        $HeaderRow = 1,
        $ParameterNameColumn = 'C',
        $HeaderMatcher
    )
    $Worksheet = Get-MZExcelWorksheet -ExcelPath $ExcelPath -WorksheetName $WorksheetName
    if ($null -eq $Worksheet ) {
        return $null
    }

    $ParamNameCol = [int][char]($ParameterNameColumn.ToUpper()) - 64
    #Get Test data set Column Array
    $CurrentCol = $ParamNameCol + 1
    $ColumnNumArray = @()
    while (-not [String]::IsNullOrEmpty($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text)) {
        if ($HeaderMatcher) {
            if ($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text -match $HeaderMatcher) {
                $ColumnNumArray += $CurrentCol
            }
        }
        else {
            $ColumnNumArray += $CurrentCol
        }
        $CurrentCol++
    }

    #Get Parameter Row Array
    $RowNumArray = @()
    $CurrentRow = $HeaderRow + 1
    $ContinuousBlankCount = 0
    do {
        if ([String]::IsNullOrEmpty($Worksheet.Cells.Item($CurrentRow, $ParamNameCol).Text)) {
            $ContinuousBlankCount++
        }
        else {
            $ContinuousBlankCount = 0
            if ("NA" -ne $Worksheet.Cells.Item($CurrentRow, $ParamNameCol).Text) {
                $RowNumArray += $CurrentRow
            }
        }
        $CurrentRow++
    }while ($ContinuousBlankCount -le $MaxBlankThreshold)

    $List = [System.Collections.ArrayList]::new()
    foreach ($iCol in $ColumnNumArray) {
        $DataSet = [ordered]@{}
        #Put Header
        $DataSet["Header"] = $Worksheet.Cells.Item($HeaderRow, $iCol).Text.Trim()
        foreach ($iRow in $RowNumArray) {
            $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim()] = $Worksheet.Cells.Item($iRow, $iCol).Text
        }
        [void]$List.Add($DataSet)
    }
    Close-MZExcelWorksheet | Out-Null
    return $List
}