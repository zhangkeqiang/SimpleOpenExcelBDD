package com.simplopen.excelbdd;
import java.io.File;

public class TestWizard {
    public static String getExcelBDDStartPath() {
        String absolutePath = new File(".").getAbsolutePath();
        String startPath = absolutePath.substring(0, absolutePath.lastIndexOf("JavaExcelBDD"));
        return startPath;
    }
}
