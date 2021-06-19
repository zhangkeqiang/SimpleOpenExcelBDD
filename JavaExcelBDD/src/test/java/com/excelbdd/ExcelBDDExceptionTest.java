package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class ExcelBDDExceptionTest {

	static Stream<Map<String, String>> provideWrongExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithExpected(filePath, "Exceptions", 2, 'C');
		return list.stream();
	}

	@ParameterizedTest(name = "#{index}-TestException: {0}")
	@MethodSource("provideWrongExampleList")
	void testGetExampleListStringString(Map<String, String> mapParams) {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
				+ mapParams.get("ExcelFileName");
		int headerRow = Behavior.getInt(mapParams.get("HeaderRow"));
		char parameterNameColumn = mapParams.get("ParameterNameColumn").charAt(0);
		Throwable exception = assertThrows(IOException.class, () -> {
			List<Map<String, String>> targetlist = Behavior.getExampleListWithExpected(filepath,
					mapParams.get("SheetName"), headerRow, parameterNameColumn);
		});
		
		System.out.println(exception.toString());
		System.out.println(exception.getClass().getSimpleName());
		assertEquals(mapParams.get("ExcelFileNameExpected"), exception.getClass().getSimpleName());
		assertTrue(exception.getMessage().indexOf(mapParams.get("SheetNameExpected")) >= 0);
	}
}
