package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class DataTableTest {

	static Stream<Map<String, String>> provideExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/DataTableBDD.xlsx";
		return Behavior.getExampleStream(filePath, "DataTableBDD", "Scenario");
	}

	@ParameterizedTest(name = "Test{index}:{0}")
	@MethodSource("provideExampleList")
	void test(Map<String, String> parameterMap) throws IOException {
		TestWizard.showMap(parameterMap);
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
				+ parameterMap.get("ExcelFileName");
		int headerRow = Double.valueOf(parameterMap.get("HeaderRow")).intValue();
		char startColumn = parameterMap.get("StartColumn").charAt(0);
		List<Map<String, String>> dataTable = Behavior.getDataTable(filePath, parameterMap.get("SheetName"), headerRow,
				startColumn);
		assertTrue(dataTable.size() > 0);
		TestWizard.showMap(dataTable.get(0));
		assertEquals(TestWizard.getInt(parameterMap.get("TestSetCount")), dataTable.size());
		assertEquals("class java.util.HashMap", dataTable.get(0).getClass().toString());
		assertEquals(parameterMap.get("FirstGridValue"), dataTable.get(0).get("Header01"));
		assertEquals(parameterMap.get("LastGridValue"), dataTable.get(5).get("Header08"));
		assertEquals(TestWizard.getInt(parameterMap.get("ColumnCount")), dataTable.get(5).size());
		// # one check is added for V0.5
		assertEquals(parameterMap.get("Header03InThirdSet"), dataTable.get(2).get("Header03"));
	}
}
