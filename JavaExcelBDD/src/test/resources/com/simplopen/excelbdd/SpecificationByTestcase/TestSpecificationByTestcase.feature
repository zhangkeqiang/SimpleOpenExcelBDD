Feature: Get Test data Set List from Excel Specification by Testcase
    Specification and its testcase are put in excel, so they should be get as a List

    Scenario Outline: Get list from a simple sheet
        Given The Excel file is "src\\test\\resources\\ExcelBDD.xlsx"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"

        When invoke get test data from excel
        Then a testset list is got, which count is <ListCount>

        Examples:
            | SheetName | HeaderRow | ParameterColumn | ListCount |
            | SBTSheet1 | 2         | B               | 4         |

