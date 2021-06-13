package com.excelbdd;

import static org.junit.Assert.*;
import static org.junit.jupiter.api.Assertions.assertEquals;


import java.util.Collection;
import java.util.Map;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

import com.excelbdd.ZMExcel;



@RunWith(Parameterized.class)
public class ZMExcelParameterizedTest {
	
	Map<String, String> mapParams;
	
	@Parameters
    public static Collection<Object[]> prepareData()
    {
    	String filepath = TestWizard.getExcelBDDStartPath() + "BDDExcel/ExcelBDD.xlsx";
    	return ZMExcel.getExampleCollection(filepath, "SimpleOpenBDD", 1, 'D');
    }
	
	public ZMExcelParameterizedTest(Map<String, String> map){
		this.mapParams = map;
	}

	@Test
	public void testZMExcelParameterizedTest() {
		
		System.out.println("Header " + mapParams.get("Header"));
		System.out.println("SheetName " + mapParams.get("SheetName"));
		System.out.println("HeaderRow " + mapParams.get("HeaderRow"));

		System.out.println(mapParams.get("MaxBlankThreshold"));
		System.out.println(mapParams.get("HeaderMatcher"));
		System.out.println(mapParams.get("ParameterCount"));		
		System.out.println("ParameterNameColumn " + mapParams.get("ParameterNameColumn"));
		
		assertEquals("Scenario1", mapParams.get("Header1Name"));
		assertEquals("V1.1", mapParams.get("FirstGridValue"));
		assertEquals("4.4", mapParams.get("LastGridValue"));
		assertEquals("V1.2", mapParams.get("ParamName1InSet2Value"));
		assertEquals("V2.2", mapParams.get("ParamName2InSet2Value"));
		assertEquals("", mapParams.get("ParamName3Value"));
	}
}
