package com.simplopen.excelbdd.SpecificationByTestcase;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;
import java.util.Map;

import com.simplopen.excelbdd.ZMExcel;

import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import io.cucumber.java.en.When;

public class FeatureSteps {
    String excelFilePath;
    String sheetName;
    int headerRow;
    String parameterColumn;
    List<Map<String, String>> list;
    @Given("The Excel file is {string}")
    public void the_excel_file_is(String string) {
        excelFilePath = string;
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

    @When("invoke get test data from excel")
    public void invoke_get_test_data_from_excel() {
        list = ZMExcel.getMZExampleWithTestResultList(excelFilePath, sheetName, headerRow, parameterColumn);
    }

    @Then("a testset list is got, which count is {int}")
    public void a_testset_list_is_got_which_count_is(Integer int1) {
        assertEquals(int1, list.size());
    }
}
