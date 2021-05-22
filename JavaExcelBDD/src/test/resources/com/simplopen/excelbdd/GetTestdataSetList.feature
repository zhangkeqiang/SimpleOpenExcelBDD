Feature: Get Test data Set List
    test data is put in excel, so they should be get as a List

    Scenario Outline: Get list from a simple sheet
        Given The Excel file is "<ExcelFilePath>"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"

        When Get the test dataset list
        Then Test dataset list which contains 5 sets is got
        And The No. 1 element of list Key "Header", Value is "Scenario1"

        Examples:
            | ExcelFilePath                       | SheetName | HeaderRow | ParameterColumn |
            | src\\test\\resources\\ExcelBDD.xlsx | Sheet1    | 1         | B               |
            | src\\test\\resources\\ExcelBDD.xlsx | Sheet2    | 1         | C               |

