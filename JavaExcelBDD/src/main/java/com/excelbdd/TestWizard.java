package com.excelbdd;
import java.io.File;

public class TestWizard {
    public static String getExcelBDDStartPath(String childPath) {
        String absolutePath = new File(".").getAbsolutePath();
        String startPath = absolutePath.substring(0, absolutePath.lastIndexOf(childPath));
        return startPath;
    }
}
