Feature: Get Test data Set List from Excel Specification by Testcase
    Specification and its testcase are put in excel, so they should be get as a List

    Scenario Outline: Get list from a simple sheet
        Given The Excel file is "src\\test\\resources\\ExcelBDD.xlsx"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"

        When invoke get test data from excel
        Then a testset list is got, which count is <ListCount>
        And The Header of 1st set is "Scenario1"
        And Input value of Variable "ParamName1" is "V1.1"
        And Expected value of Variable "ParamName1" is "V1.1"
        And Test Result value of Variable "ParamName1" is "pass"
        And The 1st data table is:
            | ParameterName | Input     | Expected  | TestResult |
            | ParamName1    | V1.1      | V1.1      | pass       |
            | ParamName2    | V2.1      | V2.1      | pass       |
            | ParamName3    |           |           | pass       |
            | ParamName4    | 2021/4/30 | 2021/4/30 | pass       |

        Examples:
            | SheetName | HeaderRow | ParameterColumn | ListCount |
            | SBTSheet1 | 2         | B               | 4         |
            | SBTSheet2 | 2         | B               | 4         |

    Scenario: excel file does not exist
        Given The Excel file is "src\\test\\resources\\NoExcelBDD.xlsx"
        When invoke on a wrong file
        Then get blank list because the file doesn't exist


    Scenario: sheet does not exist
        Given The Excel file is "src\\test\\resources\\ExcelBDD.xlsx"
        And The Sheet name is "wrongsheet"
        When invoke on a wrong sheet
        Then get blank list because the file doesn't exist