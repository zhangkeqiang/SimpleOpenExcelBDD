package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import static org.junit.jupiter.api.Assertions.*;

class ExcelBDDExceptionTest {

	static Stream<Map<String, String>> provideWrongExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithExpected(filePath, "Exceptions", 2, 'C');
		return list.stream();
	}

	@ParameterizedTest(name = "#{index}-TestException: {0}")
	@MethodSource("provideWrongExampleList")
	void testGetExampleListStringString(Map<String, String> mapParams) {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/" + mapParams.get("ExcelFileName");
		int headerRow = Behavior.getInt(mapParams.get("HeaderRow"));
		char parameterNameColumn = mapParams.get("ParameterNameColumn").charAt(0);
		List<Map<String, String>> targetlist;
		try {
			targetlist = Behavior.getExampleListWithExpected(filepath,
					mapParams.get("SheetName"), headerRow, parameterNameColumn);
			assertNull(targetlist);
		} catch (Exception e) {
			System.out.println(e.toString());
			System.out.println(e.getClass().getSimpleName());
			System.out.println(e.getMessage());
			assertEquals(mapParams.get("ExcelFileNameExpected"), e.getClass().getSimpleName());
			assertTrue(e.getMessage().indexOf(mapParams.get("SheetNameExpected")) >= 0);
		}
		
	}
}
