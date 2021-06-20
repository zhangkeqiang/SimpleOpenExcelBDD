package com.excelbdd;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Behavior {
	private static final String SIMPLE = "SIMPLE";
	private static final String TESTRESULT = "TESTRESULT";
	private static final String EXPECTED = "EXPECTED";
	private static final String ANY_MATCHER = ".*";
	public static final String NEVER_MATCHED_STRING = "i_m_p_o_s_i_b_l_e";

	private Behavior() {
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName) throws IOException {
		return getExampleList(excelPath, worksheetName, 1, 'C', ANY_MATCHER);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName)
			throws IOException {
		return getExampleList(excelPath, worksheetName).stream();
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, ANY_MATCHER,
				NEVER_MATCHED_STRING, SIMPLE);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, ANY_MATCHER,
				NEVER_MATCHED_STRING, SIMPLE).stream();
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				NEVER_MATCHED_STRING);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher).stream();
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnMatcher,
				SIMPLE);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnMatcher)
				.stream();
	}

	private static HashMap<Integer, Integer> getHeaderMap(String headerMatcher, String headerUnMatcher,
			ArrayList<Map<String, String>> listTestSet, int parameterNameColumnNum, XSSFRow rowHeader, int step) {
		// Get Matched Column HashMap
		String strRealHeaderMatcher = ANY_MATCHER + headerMatcher + ANY_MATCHER;
		String strRealHeaderUnMatcher = ANY_MATCHER + headerUnMatcher + ANY_MATCHER;
		int nMaxColumn = rowHeader.getLastCellNum();
		HashMap<Integer, Integer> mapTestSetHeader = new HashMap<>();
		int nTestSet = 0;
		for (int iCol = parameterNameColumnNum + 1; iCol < nMaxColumn; iCol += step) {
			XSSFCell cellHeader = rowHeader.getCell(iCol);
			String strHeader = cellHeader.getStringCellValue();
			if ((strHeader != null) && (!strHeader.isEmpty()) && strHeader.matches(strRealHeaderMatcher)
					&& (!strHeader.matches(strRealHeaderUnMatcher))) {
				mapTestSetHeader.put(iCol, nTestSet);
				Map<String, String> mapTestSet = new HashMap<>();
				mapTestSet.put("Header", cellHeader.getStringCellValue());
				listTestSet.add(mapTestSet);
				nTestSet++;
			}
		}
		return mapTestSetHeader;
	}

	/**
	 * @param parameterStartRow
	 * @param parameterNameColumnNum
	 * @param sheetTestData
	 * @return
	 */
	private static HashMap<Integer, String> getParameterNameMap(int parameterStartRow, int parameterNameColumnNum,
			XSSFSheet sheetTestData) {
		HashMap<Integer, String> mapParameterName = new HashMap<>();
		int nContinuousBlankCount = 0;
		for (int iRow = parameterStartRow; iRow <= sheetTestData.getLastRowNum(); iRow++) {
			if (nContinuousBlankCount > 3) {
				break;
			}
			XSSFRow rowCurrent = sheetTestData.getRow(iRow);
			if (rowCurrent == null) {
				nContinuousBlankCount++;
				continue;
			}
			XSSFCell cellParameterName = rowCurrent.getCell(parameterNameColumnNum);
			if (cellParameterName == null) {
				nContinuousBlankCount++;
				continue;
			}
			String strParameterName = cellParameterName.getStringCellValue();
			if (strParameterName == null || strParameterName.isEmpty()) {
				nContinuousBlankCount++;
			} else if (strParameterName.equals("NA")) {
				nContinuousBlankCount = 0;
			} else {
				mapParameterName.put(iRow, strParameterName);
				nContinuousBlankCount = 0;
			}
		}
		return mapParameterName;
	}

	public static Collection<Object[]> getExampleCollection(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) throws IOException {
		Collection<Object[]> collectionTestData = new ArrayList<>();
		List<Map<String, String>> listTestData = getExampleList(excelPath, worksheetName, headerRow,
				parameterNameColumn, ANY_MATCHER, NEVER_MATCHED_STRING, SIMPLE);
		for (Map<String, String> map : listTestData) {
			Object[] arrayObj = { map };
			collectionTestData.add(arrayObj);
		}
		return collectionTestData;
	}

	public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, ANY_MATCHER,
				NEVER_MATCHED_STRING, EXPECTED);
	}

	public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				NEVER_MATCHED_STRING, EXPECTED);
	}

	public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher, String headerUnMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnMatcher,
				EXPECTED);
	}

	public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn) throws IOException {
		return getExampleListWithTestResult(excelPath, worksheetName, headerRow, parameterNameColumn, ANY_MATCHER);
	}

	public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				NEVER_MATCHED_STRING, TESTRESULT);
	}

	public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher, String headerUnMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnMatcher,
				TESTRESULT);
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnMatcher, String type) throws IOException {
		// poi get row from 0, so 1st headerRow is at 0
		// by default, actualHeaderRow is below
		int actualHeaderRow = headerRow - 1;
		int actualParameterStartRow = headerRow;
		int columnStep = 1;
		if (TESTRESULT.equals(type)) {
			// because of input/expected/testresult row, the below -2
			actualParameterStartRow = headerRow + 1;
			columnStep = 3;
		} else if (EXPECTED.equals(type)) {
			actualParameterStartRow = headerRow + 1;
			columnStep = 2;
		}
		ArrayList<Map<String, String>> listTestSet = new ArrayList<>();
		// poi get column from 0, so Column A's Num is 0, 65 is A's ASCII code
		int parameterNameColumnNum = (int) parameterNameColumn - 65;

		FileInputStream excelFile = new FileInputStream(new File(excelPath));
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		XSSFSheet sheetTestData = workbook.getSheet(worksheetName);
		if (sheetTestData == null) {
			throw new IOException(worksheetName + " does not exist.");
		}

		XSSFRow rowHeader = sheetTestData.getRow(actualHeaderRow);
		HashMap<Integer, Integer> mapTestSetHeader = getHeaderMap(headerMatcher, headerUnMatcher, listTestSet,
				parameterNameColumnNum, rowHeader, columnStep);

		// Get ParameterNames HashMap
		HashMap<Integer, String> mapParameterName = getParameterNameMap(actualParameterStartRow, parameterNameColumnNum,
				sheetTestData);

		for (Map.Entry<Integer, String> aParameterName : mapParameterName.entrySet()) {
			int iRow = aParameterName.getKey();
			String strParameterName = aParameterName.getValue();
			XSSFRow rowCurrent = sheetTestData.getRow(iRow);

			for (Map.Entry<Integer, Integer> entryHeader : mapTestSetHeader.entrySet()) {
				int iCol = entryHeader.getKey();
				Map<String, String> mapTestSet = listTestSet.get(entryHeader.getValue());
				putParameter(strParameterName, rowCurrent, mapTestSet, iCol);
				if (columnStep > 1) {
					putParameter(strParameterName + "Expected", rowCurrent, mapTestSet, iCol + 1);
					if (columnStep == 3) {
						putParameter(strParameterName + "TestResult", rowCurrent, mapTestSet, iCol + 2);
					}
				}
			}
		}
		return listTestSet;
	}

	/**
	 * @param strParameterName
	 * @param rowCurrent
	 * @param mapTestSet
	 * @param iCol
	 */
	private static void putParameter(String strParameterName, XSSFRow rowCurrent, Map<String, String> mapTestSet,
			int iCol) {

		XSSFCell cellCurrent = rowCurrent.getCell(iCol);
		if (cellCurrent.getCellType() == CellType.STRING) {
			mapTestSet.put(strParameterName, cellCurrent.getStringCellValue());
		} else if (cellCurrent.getCellType() == CellType.NUMERIC) {
			mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getNumericCellValue()));
		} else if (cellCurrent.getCellType() == CellType._NONE) {
			mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getDateCellValue()));
		} else if (cellCurrent.getCellType() == CellType.BLANK) {
			mapTestSet.put(strParameterName, "");
		} else if (cellCurrent.getCellType() == CellType.BOOLEAN) {
			mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getBooleanCellValue()));
		} else if (cellCurrent.getCellType() == CellType.FORMULA) {
			mapTestSet.put(strParameterName, cellCurrent.getRawValue());
		} else {
			mapTestSet.put(strParameterName, cellCurrent.getRawValue());
		}
	}
}
