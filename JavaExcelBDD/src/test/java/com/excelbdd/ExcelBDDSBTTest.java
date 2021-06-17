/**
 * 
 */
package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

/**
 * @author Mike
 *
 */
class ExcelBDDSBTTest {

	/**
	 * @throws java.lang.Exception
	 */
	@BeforeAll
	static void setUpBeforeClass() throws Exception {
	}

	/**
	 * @throws java.lang.Exception
	 */
	@AfterAll
	static void tearDownAfterClass() throws Exception {
	}

	/**
	 * @throws java.lang.Exception
	 */
	@BeforeEach
	void setUp() throws Exception {
	}

	/**
	 * @throws java.lang.Exception
	 */
	@AfterEach
	void tearDown() throws Exception {
	}

	/**
	 * Test method for {@link com.excelbdd.Behavior#getExampleListWithTestResult(java.lang.String, java.lang.String, int, char)}.
	 */
	@Test
	void testGetExampleListWithTestResultStringStringIntChar() {
//		fail("Not yet implemented");
	}
	
	
	static Stream<Map<String, String>> provideExampleListWithExpected() {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithExpected(filepath, "Expected1", 2, 'B');
		return list.stream();
	}

	/**
	 * Test method for {@link com.excelbdd.Behavior#getExampleListWithExpected(java.lang.String, java.lang.String, int, char)}.
	 */
	@ParameterizedTest(name = "#{index}-TestExpected : {0}")
	@MethodSource("provideExampleListWithExpected")
	void testGetExampleListWithExpected(Map<String, String> mapParams) {
		assertNotNull(mapParams);
		System.out.println("Header " + mapParams.get("Header"));
		System.out.println(String.format("ParamName1 %s", mapParams.get("ParamName1")));
		System.out.println(String.format("ParamName1 %s", mapParams.get("ParamName2")));
		System.out.println("ParamName3 " + mapParams.get("ParamName3"));
		assertTrue(mapParams.get("ParamName1").startsWith("V1."));
		assertTrue(mapParams.get("ParamName2").startsWith("V2."));
	}

	/**
	 * Test method for {@link com.excelbdd.Behavior#getExampleListWithTestResult(java.lang.String, java.lang.String, int, char, java.lang.String)}.
	 */
	@Test
	void testGetExampleListWithTestResultStringStringIntCharString() {
//		fail("Not yet implemented");
	}

}
