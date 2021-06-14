package com.excelbdd;

import java.io.File;

public class TestWizard {
	private TestWizard() {
	}

	public static String getExcelBDDStartPath(String childPath) {
		String absolutePath = new File(".").getAbsolutePath();
		return absolutePath.substring(0, absolutePath.lastIndexOf(childPath));
	}
}
