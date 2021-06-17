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
		TestWizard.showMap(mapParams);
		assertTrue(mapParams.get("ParamName1").startsWith("V1."));
		assertTrue(mapParams.get("ParamName2").startsWith("V2."));
		assertEquals(true,mapParams.get("ParamName3").isEmpty());
		assertEquals(false,mapParams.get("ParamName4").isEmpty());
		
		assertTrue(mapParams.get("ParamName1Expected").startsWith("V1."));
		assertTrue(mapParams.get("ParamName2Expected").startsWith("V2."));
		assertEquals(true,mapParams.get("ParamName3Expected").isEmpty());
		assertEquals(false,mapParams.get("ParamName4Expected").isEmpty());
		
		assertNull(mapParams.get("ParamName1TestResult"));
		assertNull(mapParams.get("ParamName2TestResult"));
		assertNull(mapParams.get("ParamName3TestResult"));
		assertNull(mapParams.get("ParamName4TestResult"));
	}

	static Stream<Map<String, String>> provideExampleListWithTestResultStringStringIntCharString() {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithTestResult(filepath, "SBTSheet1", 2, 'B',"5");
		return list.stream();
	}
	/**
	 * Test method for {@link com.excelbdd.Behavior#getExampleListWithTestResult(java.lang.String, java.lang.String, int, char, java.lang.String)}.
	 */
	@ParameterizedTest(name = "TestResult{index}:{0}")
	@MethodSource("provideExampleListWithTestResultStringStringIntCharString")
	void testGetExampleListWithTestResultStringStringIntCharString(Map<String, String> mapParams) {
		assertNotNull(mapParams);
		TestWizard.showMap(mapParams);
		assertTrue(mapParams.get("ParamName1").startsWith("V1."));
		assertTrue(mapParams.get("ParamName2").startsWith("V2."));
		assertEquals(true,mapParams.get("ParamName3").isEmpty());
		assertEquals(false,mapParams.get("ParamName4").isEmpty());
		
		assertTrue(mapParams.get("ParamName1Expected").startsWith("V1."));
		assertTrue(mapParams.get("ParamName2Expected").startsWith("V2."));
		assertEquals(true,mapParams.get("ParamName3Expected").isEmpty());
		assertEquals(false,mapParams.get("ParamName4Expected").isEmpty());
		
		assertNotNull(mapParams.get("ParamName1TestResult"));
		assertNotNull(mapParams.get("ParamName2TestResult"));
		assertNotNull(mapParams.get("ParamName3TestResult"));
		assertNotNull(mapParams.get("ParamName4TestResult"));
		
		assertEquals("pass",mapParams.get("ParamName1TestResult"));
		assertEquals("pass",mapParams.get("ParamName2TestResult"));
		assertEquals("pass",mapParams.get("ParamName3TestResult"));
		assertEquals("pass",mapParams.get("ParamName4TestResult"));
	}

}
