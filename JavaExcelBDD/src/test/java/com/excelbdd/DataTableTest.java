package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class DataTableTest {

	static Stream<Map<String, String>> provideExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/DataTableBDD.xlsx";
		return Behavior.getExampleStream(filePath, "DataTableBDD");
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
		// $ExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
		// $DataTableA = Get-DataTable -ExcelPath $ExcelPath -WorksheetName $SheetName `
		// -HeaderRow $HeaderRow -StartColumn $StartColumn
		//
		// Show-ExampleList $DataTableA
		// $DataTableA.Count | Should -Be $TestSetCount
		// $DataTableA.GetType().Name | Should -Be 'Object[]'
		// $DataTableA[0].GetType().Name | Should -Be 'HashTable'
		// $DataTableA[0]["Header01"] | Should -Be $FirstGridValue
		// $DataTableA[5]["Header08"] | Should -Be $LastGridValue
		// $DataTableA[5].Count | Should -Be $ColumnCount
		// # one check is added for V0.5
		// $DataTableA[2]["Header03"] | Should -Be $Header03InThirdSet

	}
}
