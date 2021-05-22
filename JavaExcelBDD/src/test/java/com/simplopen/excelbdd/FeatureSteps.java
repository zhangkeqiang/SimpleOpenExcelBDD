package com.simplopen.excelbdd;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;
import java.util.Map;

import io.cucumber.java.en.*;

public class FeatureSteps {
    String excelFilePath;
    String sheetName;
    int headerRow;
    char parameterColumn;
    List<Map> list;

    @Given("The Excel file is {string}")
    public void the_excel_file_is(String excelFile) {
        excelFilePath = excelFile;
    }

    @Given("The Sheet name is {string}")
    public void the_sheet_name_is(String sheetName) {
        this.sheetName = sheetName;
    }

    @Given("Header Row is {int}")
    public void header_row_is(Integer int1) {
        headerRow = int1;
    }

    @Given("Parameter Column is {string}")
    public void parameter_column_is(String string) {
        parameterColumn = string.charAt(0);
    }

    @When("Get the test dataset list")
    public void get_the_test_dataset_list() {
        list = ZMExcel.getExampleList(excelFilePath, sheetName, headerRow, parameterColumn);
    }

    @Then("Test dataset list which contains {int} sets is got")
    public void test_dataset_list_which_contains_sets_is_got(Integer int1) {
        assertEquals(int1, list.size());
    }

    @Then("The No. {int} element of list Key {string}, Value is {string}")
    public void the_no_element_of_list_key_value_is(Integer int1, String key, String value) {
        assertEquals(value, list.get(int1-1).get(key));
    }
}
