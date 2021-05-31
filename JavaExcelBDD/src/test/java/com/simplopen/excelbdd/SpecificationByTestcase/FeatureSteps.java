package com.simplopen.excelbdd.SpecificationByTestcase;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

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
	char parameterNameColumn;
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
		parameterNameColumn = string.charAt(0);
	}

	@When("invoke get test data from excel")
	public void invoke_get_test_data_from_excel() {
		list = ZMExcel.getMZExampleWithTestResultList(excelFilePath, sheetName, headerRow, parameterNameColumn);
	}

	@Then("a testset list is got, which count is {int}")
	public void a_testset_list_is_got_which_count_is(int int1) {
		assertEquals(int1, list.size());
	}

	@Then("The Header of 1st set is {string}")
	public void the_header_of_1st_set_is(String string) {
		assertEquals(string, list.get(0).get("Header"));
	}

	@Then("Input value of Variable {string} is {string}")
	public void input_value_of_variable_is(String string, String string2) {
		assertEquals(string2, list.get(0).get(string));
	}

	@Then("Expected value of Variable {string} is {string}")
	public void expected_value_of_variable_is(String string, String string2) {
		assertEquals(string2, list.get(0).get(string + "Expected"));
	}

	@Then("Test Result value of Variable {string} is {string}")
	public void test_result_value_of_variable_is(String string, String string2) {
		assertEquals(string2, list.get(0).get(string + "TestResult"));
	}

	@Then("The 1st data table is:")
	public void the_data_table_is(io.cucumber.datatable.DataTable dataTable) {
		List<Map<String, String>> mapList = dataTable.asMaps();
		int i = 0;
		for (Map<String, String> map : mapList) {
			System.out.println("===========");
			for (Map.Entry<String, String> mapEntry : map.entrySet()) {
				System.out.print(mapEntry.getKey() + " --- ");
				System.out.println(mapEntry.getValue());
			}
			if (map.get("Input") == null) {
				assertEquals("", list.get(i).get(map.get("ParameterName")));
			} else {
				assertEquals(map.get("Input"), list.get(i).get(map.get("ParameterName")));
			}

			if (map.get("Expected") == null) {
				assertEquals("", list.get(i).get(map.get("ParameterName") + "Expected"));
			} else {
				assertEquals(map.get("Expected"), list.get(i).get(map.get("ParameterName") + "Expected"));
			}
			assertEquals(map.get("TestResult"), "pass");
			assertEquals(map.get("TestResult"), list.get(i).get(map.get("ParameterName") + "TestResult"));

		}

	}

	@When("invoke on a wrong file")
	public void invoke_on_a_wrong_file() {
		list = ZMExcel.getMZExampleWithTestResultList(excelFilePath, "sheetName", 1, 'B');
	}

	@Then("get blank list because the file doesn't exist")
	public void get_blank_list_because_the_file_doesn_t_exist() {
		assertNotNull(list);
		assertEquals(0, list.size());
	}

	@When("invoke on a wrong sheet")
	public void invoke_on_a_wrong_sheet() {
		list = ZMExcel.getMZExampleWithTestResultList(excelFilePath, sheetName, 1, 'B');
	}
}
