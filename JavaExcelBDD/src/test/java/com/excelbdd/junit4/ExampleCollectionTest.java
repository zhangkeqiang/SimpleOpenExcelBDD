package com.excelbdd.junit4;

import static org.junit.Assert.*;

import java.io.IOException;
import java.util.Collection;
import java.util.Map;

import org.junit.Test;

import com.excelbdd.Behavior;
import com.excelbdd.TestWizard;

public class ExampleCollectionTest {

	@Test
	public void testMatcher() throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		Collection<Object[]> exampleCollection = Behavior.getExampleCollection(filepath, "SpecificationByExample", "Scenario1");
		assertEquals(2,exampleCollection.size());
	}
	
	@Test
	public void testMatcherUnMatcher() throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		Collection<Object[]> exampleCollection = Behavior.getExampleCollection(filepath, "SpecificationByExample", "Scenario1","easy");
		assertEquals(1,exampleCollection.size());
		Object[] ObjectList = exampleCollection.toArray();
		Object[] ob = (Object[]) ObjectList[0];
		System.out.println(ob[0]);
		TestWizard.showMap((Map<String, String>) ob[0]);
	}

}
