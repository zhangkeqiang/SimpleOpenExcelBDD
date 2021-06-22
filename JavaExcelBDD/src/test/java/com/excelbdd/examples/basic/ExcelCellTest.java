package com.excelbdd.examples.basic;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import com.excelbdd.Behavior;
import com.excelbdd.TestWizard;

class ExcelCellTest {

	static Stream<Map<String, String>> provideExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDDSampleA.xlsx";
		return Behavior.getExampleStream(filePath, "CellValue");
	}

	@ParameterizedTest(name = "Test{index}:{0}")
	@MethodSource("provideExampleList")
	void testgetCellValue(Map<String, String> mapParams) throws IOException {
		TestWizard.showMap(mapParams);
		String strCellValue = mapParams.get("CellValue");
		switch (mapParams.get("CellValueExpected")) {
			case "numberic":
				try {
					double d = Double.parseDouble(strCellValue);
				} catch (NumberFormatException e) {
					fail("it is not numberic.");
				}
				break;
			case "string":
				assertfalse(strCellValue.isEmpty());
				break;
			case "boolean":
				assertTrue(strCellValue.matches("true|false"));
				break;
			case "blank":
				assertTrue(strCellValue.isEmpty());
				break;
			default:
				fail("no others");
		}

	}

	private void assertfalse(boolean empty) {
		// TODO Auto-generated method stub

	}

}
