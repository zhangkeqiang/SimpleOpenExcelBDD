package com.excelbdd;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import static org.junit.jupiter.api.Assertions.*;

public class ExcelBDDSBETest {

	static Stream<Map<String, String>> provideExampleList() {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleList(filepath, "SpecificationByExample", 1, 'E');
		return list.stream();
	}

	@ParameterizedTest(name = "#{index} - Test with Map : {0}")
	@MethodSource("provideExampleList")
	void testParameterizedTestByMap(Map<String, String> mapParams) {
		assertNotNull(mapParams);
		System.out.println("Header " + mapParams.get("Header"));
		System.out.println("SheetName " + mapParams.get("SheetName"));
		System.out.println("HeaderRow " + mapParams.get("HeaderRow"));
		System.out.println("ParameterNameColumn " + mapParams.get("ParameterNameColumn"));
		assertEquals("Scenario1", mapParams.get("Header1Name"));
		assertEquals("V1.1", mapParams.get("FirstGridValue"));
		assertEquals("4.4", mapParams.get("LastGridValue"));
		assertEquals("V1.2", mapParams.get("ParamName1InSet2Value"));
		assertEquals("V2.2", mapParams.get("ParamName2InSet2Value"));
		assertEquals("", mapParams.get("ParamName3Value"));

		assertEquals("3.0", mapParams.get("MaxBlankThreshold"));
		System.out.println("HeaderMatcher " + mapParams.get("HeaderMatcher"));
		assertEquals("Scenario", mapParams.get("HeaderMatcher"));
		

		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		int nHeaderRow = Double.valueOf(mapParams.get("HeaderRow")).intValue();
		char charParameterNameColumn = mapParams.get("ParameterNameColumn").charAt(0);
		System.out.println("ParameterNameColumn " + charParameterNameColumn);

		List<Map<String, String>> list = Behavior.getExampleList(filepath, mapParams.get("SheetName"), nHeaderRow,
				charParameterNameColumn, mapParams.get("HeaderMatcher"));

		System.out.println("ParameterCount " + mapParams.get("ParameterCount"));
		assertEquals(mapParams.get("ParameterCount"),list.get(0).size());
		System.out.println(list.get(0).toString());
		System.out.println(list.get(1).toString());
		System.out.println(list.get(2).toString());
		System.out.println(list.get(3).toString());

		// int testDataSetCount =
		// Double.valueOf(mapParams.get("TestDataSetCount")).intValue();
		int testDataSetCount = Behavior.getInt(mapParams.get("TestDataSetCount"));
		assertEquals(testDataSetCount, list.size());

		assertEquals("V1.1", list.get(0).get("ParamName1"));
		assertEquals("V1.2", list.get(1).get("ParamName1"));
		assertEquals("V1.3", list.get(2).get("ParamName1"));
		assertEquals("V1.4", list.get(3).get("ParamName1"));

		assertEquals("V2.1", list.get(0).get("ParamName2"));
		assertEquals("V2.2", list.get(1).get("ParamName2"));

		assertEquals("", list.get(0).get("ParamName3"));
		assertEquals("", list.get(1).get("ParamName3"));
		assertEquals("", list.get(2).get("ParamName3"));
		assertEquals("", list.get(3).get("ParamName3"));

		assertEquals("2021/4/30", list.get(0).get("ParamName4"));
		assertEquals("false", list.get(1).get("ParamName4"));
		assertEquals("true", list.get(2).get("ParamName4"));
		assertEquals("4.4", list.get(3).get("ParamName4"));
	}

	@Test
	void testgetInt() {
		assertEquals(5, Behavior.getInt("5.5666"));
		assertEquals(5, Behavior.getInt("5"));
		assertEquals(5, Behavior.getInt("5.99999"));
	}

	@Test
	void testBDDExcelPath2() {
		String ExcelFilePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		File f = new File(ExcelFilePath);
		assertTrue(f.exists());
	}
}
