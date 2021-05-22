package simpleOpenAgile.BDD;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
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

public class ZMExcel {

	@SuppressWarnings("rawtypes")
	public static List<Map> getExampleList(String excelPath, String worksheetName) {
		return getExampleList(excelPath, worksheetName, 1, 'C', "");
	}

	@SuppressWarnings("rawtypes")
	public static List<Map> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, "");
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static List<Map> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher) {
		String strRealHeaderMatcher = ".*" + headerMatcher + ".*";
		ArrayList<Map> listTestSet = new ArrayList<Map>();
		FileInputStream excelFile = null;
		XSSFWorkbook workbook = null;
		try {
			int parameterNameColumnNum = (int) parameterNameColumn - 65; // poi get column from 0, so Column A's Num is
																			// 0, 65 is A's ASCII code

			excelFile = new FileInputStream(new File(excelPath));
			workbook = new XSSFWorkbook(excelFile);
			XSSFSheet sheetTestData = workbook.getSheet(worksheetName);
			XSSFRow rowHeader = sheetTestData.getRow(headerRow - 1); // poi get row from 0, so 1st headerRow is at 0

			// Get Matched Column HashMap
			int nMaxColumn = rowHeader.getLastCellNum();
			HashMap<Integer, String> mapTestDataHeader = new HashMap();
			for (int iCol = parameterNameColumnNum + 1; iCol < nMaxColumn; iCol++) {
				XSSFCell cellHeader = rowHeader.getCell(iCol);
				if (cellHeader.getStringCellValue().matches(strRealHeaderMatcher)) {
					mapTestDataHeader.put(iCol, cellHeader.getStringCellValue());
					Map<String, String> mapTestSet = new HashMap();
					mapTestSet.put("Header", cellHeader.getStringCellValue());
					listTestSet.add(mapTestSet);
				}
			}
			// Get ParameterNames HashMap
			HashMap<Integer, String> mapParameterName = new HashMap();
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
				} else if (strParameterName != "NA") {
					mapParameterName.put(iRow, strParameterName);
					nContinuousBlankCount = 0;
				} else {
					nContinuousBlankCount = 0;
				}
			}

			for (Map.Entry aParameterName : mapParameterName.entrySet()) {
				int iRow = Integer.parseInt(aParameterName.getKey().toString());
				String strParameterName = (String) aParameterName.getValue();
				XSSFRow rowCurrent = sheetTestData.getRow(iRow);
				int nPos = 0;
				for (Integer iCol : mapTestDataHeader.keySet()) {
					Map<String, String> mapTestSet = listTestSet.get(nPos++);
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

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				excelFile.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return listTestSet;
	}

	@SuppressWarnings("rawtypes")
	public static Collection<Object[]> getExampleCollection(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) {
		Collection<Object[]> collectionTestData = new ArrayList<Object[]>();
		List<Map> listTestData = getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn);
		for (Map map : listTestData) {
			Object[] arrayObj = { map };
			collectionTestData.add(arrayObj);
		}
		return collectionTestData;
	}

}
