Feature: Get Test data Set List from Excel Specification by Testcase
    Specification and its testcase are put in excel, so they should be get as a List

    Scenario Outline: Get list from a simple sheet
        Given The Excel file is "src\\test\\resources\\ExcelBDD.xlsx"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"

        Examples:
            | SheetName | HeaderRow | ParameterColumn |
            | SBTSheet1 | 2         | B               |

