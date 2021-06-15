package com.excelbdd;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class Behavior {
	protected static Logger log = LogManager.getLogger();

	private Behavior() {
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName) {
		return getExampleList(excelPath, worksheetName, 1, 'C', "");
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, "");
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher) {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				"i_m_p_o_s_i_b_l_e");
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnMatcher) {

		ArrayList<Map<String, String>> listTestSet = new ArrayList<>();
		// poi get column from 0, so Column A's Num is 0, 65 is A's ASCII code
		int parameterNameColumnNum = (int) parameterNameColumn - 65;
		try (FileInputStream excelFile = new FileInputStream(new File(excelPath));
				XSSFWorkbook workbook = new XSSFWorkbook(excelFile);) {

			XSSFSheet sheetTestData = workbook.getSheet(worksheetName);
			// poi get row from 0, so 1st headerRow is at 0
			XSSFRow rowHeader = sheetTestData.getRow(headerRow - 1);

			HashMap<Integer, Integer> mapTestSetHeader = getHeaderMap(headerMatcher, headerUnMatcher, listTestSet,
					parameterNameColumnNum, rowHeader, 1);

			// Get ParameterNames HashMap
			HashMap<Integer, String> mapParameterName = getParameterNameMap(headerRow, parameterNameColumnNum,
					sheetTestData);

			for (Map.Entry<Integer, String> aParameterName : mapParameterName.entrySet()) {
				int iRow = aParameterName.getKey();
				String strParameterName = aParameterName.getValue();
				XSSFRow rowCurrent = sheetTestData.getRow(iRow);
				for (Map.Entry<Integer, Integer> entryHeader : mapTestSetHeader.entrySet()) {
					int iCol = entryHeader.getKey();
					Map<String, String> mapTestSet = listTestSet.get(entryHeader.getValue());
					putParameter(strParameterName, rowCurrent, mapTestSet, iCol);
				}
			}
		} catch (IOException e) {
			log.error(e.getStackTrace());
		}

		return listTestSet;
	}

	private static HashMap<Integer, Integer> getHeaderMap(String headerMatcher, String headerUnMatcher,
			ArrayList<Map<String, String>> listTestSet, int parameterNameColumnNum, XSSFRow rowHeader, int step) {
		// Get Matched Column HashMap
		String strRealHeaderMatcher = ".*" + headerMatcher + ".*";
		String strRealHeaderUnMatcher = ".*" + headerUnMatcher + ".*";
		int nMaxColumn = rowHeader.getLastCellNum();
		HashMap<Integer, Integer> mapTestSetHeader = new HashMap<>();
		int nTestSet = 0;
		for (int iCol = parameterNameColumnNum + 1; iCol < nMaxColumn; iCol += step) {
			XSSFCell cellHeader = rowHeader.getCell(iCol);
			if (cellHeader.getStringCellValue().matches(strRealHeaderMatcher)) {
				if (!cellHeader.getStringCellValue().matches(strRealHeaderUnMatcher)) {
					mapTestSetHeader.put(iCol, nTestSet);
					Map<String, String> mapTestSet = new HashMap<>();
					mapTestSet.put("Header", cellHeader.getStringCellValue());
					listTestSet.add(mapTestSet);
					nTestSet++;
				}
			}
		}
		return mapTestSetHeader;
	}

	/**
	 * @param headerRow
	 * @param parameterNameColumnNum
	 * @param sheetTestData
	 * @return
	 */
	private static HashMap<Integer, String> getParameterNameMap(int headerRow, int parameterNameColumnNum,
			XSSFSheet sheetTestData) {
		HashMap<Integer, String> mapParameterName = new HashMap<>();
		int nContinuousBlankCount = 0;
		for (int iRow = headerRow; iRow <= sheetTestData.getLastRowNum(); iRow++) {
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
			char parameterNameColumn) {
		Collection<Object[]> collectionTestData = new ArrayList<>();
		List<Map<String, String>> listTestData = getExampleList(excelPath, worksheetName, headerRow,
				parameterNameColumn);
		for (Map<String, String> map : listTestData) {
			Object[] arrayObj = { map };
			collectionTestData.add(arrayObj);
		}
		return collectionTestData;
	}

	public static List<Map<String, String>> getMZExampleWithTestResultList(String excelPath, String sheetName,
			int headerRow, char parameterNameColumn) {
		String headerMatcher = ".*";
		return getMZExampleWithTestResultList(excelPath, sheetName, headerRow, headerMatcher, parameterNameColumn);

	}

	public static List<Map<String, String>> getMZExampleWithTestResultList(String excelPath, String sheetName,
			int headerRow, String headerMatcher, char parameterNameColumn) {
		ArrayList<Map<String, String>> listTestSet = new ArrayList<>();
		int parameterNameColumnNum = (int) parameterNameColumn - 65;

		try (FileInputStream excelFile = new FileInputStream(new File(excelPath));
				XSSFWorkbook workbook = new XSSFWorkbook(excelFile)) {

			XSSFSheet sheetTestData = workbook.getSheet(sheetName);
			if (sheetTestData == null) {
				log.error("%s does not exist.", sheetName);
				return listTestSet;
			}
			// poi get row from 0, so 1st headerRow is at 0
			// because of input/expected/testresult row, the below -2
			XSSFRow rowHeader = sheetTestData.getRow(headerRow - 2);
			HashMap<Integer, Integer> mapTestSetHeader = getHeaderMap(headerMatcher, "never_matched",listTestSet,
					parameterNameColumnNum, rowHeader, 3);

			// Get ParameterNames HashMap
			HashMap<Integer, String> mapParameterName = getParameterNameMap(headerRow, parameterNameColumnNum,
					sheetTestData);

			for (Map.Entry<Integer, String> aParameterName : mapParameterName.entrySet()) {
				int iRow = aParameterName.getKey();
				String strParameterName = aParameterName.getValue();
				XSSFRow rowCurrent = sheetTestData.getRow(iRow);

				for (Map.Entry<Integer, Integer> entryHeader : mapTestSetHeader.entrySet()) {
					int iCol = entryHeader.getKey();
					Map<String, String> mapTestSet = listTestSet.get(entryHeader.getValue());
					putParameter(strParameterName, rowCurrent, mapTestSet, iCol);
					putParameter(strParameterName + "Expected", rowCurrent, mapTestSet, iCol + 1);
					putParameter(strParameterName + "TestResult", rowCurrent, mapTestSet, iCol + 2);
				}
			}
		} catch (IOException e) {
			log.error(excelPath + "does not exist.");
			log.error(e.getStackTrace());
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

	public static int getInt(String string) {
		return Double.valueOf(string).intValue();
	}
}
