package com.simplopen.excelbdd.SpecificationByTestcase;

import io.cucumber.java.en.Given;

public class FeatureSteps {
    String excelFileName;
    String sheetName;
    int headerRow;
    String parameterColumn;

    @Given("The Excel file is {string}")
    public void the_excel_file_is(String string) {
        excelFileName = string;
    }

    @Given("The Sheet name is {string}")
    public void the_sheet_name_is(String string) {
        sheetName = string;
    }

    @Given("Header Row is {int}")
    public void header_row_is(Integer int1) {
        headerRow = (int) int1;
    }

    @Given("Parameter Column is {string}")
    public void parameter_column_is(String string) {
        parameterColumn = string;
    }
}
