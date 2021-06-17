package com.excelbdd;

import java.io.File;
import java.util.Map;

public class TestWizard {
	private TestWizard() {
	}

	public static String getExcelBDDStartPath(String childPath) {
		String absolutePath = new File(".").getAbsolutePath();
		return absolutePath.substring(0, absolutePath.lastIndexOf(childPath));
	}

	public static void showMap(Map<String, String> mapParams) {
		System.out.println(String.format("=======Header: %s=====", mapParams.get("Header")));
		for (Map.Entry<String, String> param : mapParams.entrySet()) {
			System.out.println(String.format("%s --- %s", param.getKey(), param.getValue()));
		}
	}
}
