package com.swapson.Util;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 
 * @author Swapnil Sonar
 *
 */
public class ExcelReader {
	// private String excelPath;
	private Workbook wb;

	public static final Map<String, String> REFERENCE_DATA = new HashMap<String, String>();

	public ExcelReader(String excelPath) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// this.excelPath = excelPath;
		wb = WorkbookFactory.create(new File(excelPath));
	}

	/**
	 * Load reference data for substituting place holder values enclosed in ${ }
	 * For Example: ${userName}
	 */
	public void loadReferenceData() {
		Sheet referenceDataSheet = wb.getSheet(Constants.REFERENCE);
		int rowCount = referenceDataSheet.getPhysicalNumberOfRows();
		for (int r = referenceDataSheet.getFirstRowNum() + 1; r < rowCount; r++) {
			Row row = referenceDataSheet.getRow(r);
			
			String refKey = (String) readCellData(row.getCell(0));
			String refValue = (String) readCellData(row.getCell(1));
			
			// format reference data here. 
			// For Example, if you want to change data format, check the place holder name & format the value.
			/*if("${userName}".equalsIgnoreCase(refKey)){
				refValue = refValue.toUpperCase();
			}*/
			REFERENCE_DATA.put(refKey, refValue);
						
		}

	}

	/**
	 * 
	 * @return returns all test cases list
	 */
	public List<TestCase> getTestCases() {
		loadReferenceData();
		List<TestCase> testCases = new ArrayList<TestCase>();
		Sheet testCaseSheet = wb.getSheet(Constants.TEST_CASES);
		int rowCount = testCaseSheet.getPhysicalNumberOfRows();
		for (int r = testCaseSheet.getFirstRowNum() + 1; r < rowCount; r++) {
			Row row = testCaseSheet.getRow(r);
			TestCase testCase = new TestCase(row.getCell(0).getStringCellValue().trim(),
					row.getCell(1).getStringCellValue().trim(), row.getCell(2).getStringCellValue().trim());
			testCases.add(testCase);
		}
		return testCases;
	}

	/**
	 * return all test data
	 * 
	 * @param includeRunModeNo
	 *            true, will include all the test cases & test data irrespective of "Run Mode" 
	 *            false, will include only the test cases & test data for which "Run Mode" is Y
	 * @return
	 */
	public Map<String, List<Map<String, String>>> readAllTestData(final boolean includeRunModeNo) {
		Map<String, List<Map<String, String>>> allTestData = new LinkedHashMap<String, List<Map<String, String>>>();
		List<TestCase> testCases = getTestCases();
		for (TestCase tc : testCases) {
			if (Constants.RUN_MODE_YES.equalsIgnoreCase(tc.getRunMode()) || includeRunModeNo) {
				allTestData.put(tc.getTestCaseName(), readTestData(tc.getTestCaseName(), includeRunModeNo));
			}
		}
		return allTestData;
	}

	/**
	 * Read test data for given test case The test data reading start after
	 * "TEST_CASE_NAME START" marker till it finds "TEST_CASE_NAME END" marker
	 * 
	 * @param testCaseName
	 *            test case name for fetching data
	 * @param includeRunModeNo
	 *            true, will include test data irrespective of "Run Mode" 
	 *            false, will include only the test data for which "Run Mode" is Y
	 * @return returns test data for given test case
	 */
	public List<Map<String, String>> readTestData(final String testCaseName, final boolean includeRunModeNo) {
		List<Map<String, String>> testData = new ArrayList<Map<String, String>>();
		Sheet testDatasheet = wb.getSheet(Constants.DATA);
		int rowCount = testDatasheet.getLastRowNum();
		List<String> headers = new ArrayList<String>();

		boolean testDataFound = false;
		for (int rowNum = testDatasheet.getFirstRowNum(); rowNum < rowCount; rowNum++) {
			Row row = testDatasheet.getRow(rowNum);
			String firstColumnValue = (row != null && row.getCell(0) != null) ? row.getCell(0).getStringCellValue()
					: "";
			if (firstColumnValue.trim().matches((testCaseName + Constants.TEST_CASE_START_MARKER))) {
				testDataFound = true;
				int testCaseRowNum = rowNum + 1;
				for (; !testDatasheet.getRow(testCaseRowNum).getCell(0).getStringCellValue()
						.matches((testCaseName + Constants.TEST_CASE_END_MARKER)) && testCaseRowNum < rowCount; testCaseRowNum++) {
					if (testCaseRowNum == rowNum + 1) {
						// first row after test cases start marker; read header
						for (int c = 0; c < testDatasheet.getRow(testCaseRowNum).getLastCellNum(); c++) {
							if (!testDatasheet.getRow(testCaseRowNum).getCell(c).getStringCellValue().equalsIgnoreCase("")) {
								headers.add(testDatasheet.getRow(testCaseRowNum).getCell(c).getStringCellValue());
							}
						}
					} else {
						Map<String, String> rowData = new LinkedHashMap<String, String>();
						if (Constants.RUN_MODE_YES.equalsIgnoreCase(
								testDatasheet.getRow(testCaseRowNum).getCell(0).getStringCellValue()) || includeRunModeNo) {
							for (int colNum = 0; colNum < testDatasheet.getRow(testCaseRowNum).getLastCellNum() && colNum < headers.size(); colNum++) {
								String value = (String) readCellData(testDatasheet.getRow(testCaseRowNum).getCell(colNum));

								Pattern re = Pattern.compile("\\$\\{([^}]+)\\}");
								Matcher m = re.matcher(value);
								if (m.groupCount() > 0) {
									while (m.find()) {
										String key = m.group();

										if (REFERENCE_DATA.containsKey(key)) {
											value = value.replace(key, REFERENCE_DATA.get(key));
										}
									}
								} 
								
								// format test data here. 
								// For Example, if you want to change particular column value format, 
								// if testCaseName & column number matches, format the value
								/*if("TestCase1".equalsIgnoreCase(testCaseName) && colNum == 2) {
									value = String.format("%.04f",Double.parseDouble(value));
								}*/
								rowData.put(headers.get(colNum), value);
								
							}
							testData.add(rowData);
						}
					}
				}
				rowNum = testCaseRowNum;
			}
			if (testDataFound) {
				break;
			}
		}

		return testData;
	}

	/**
	 * read cell data
	 * 
	 * @param cell
	 * @return
	 */
	public Object readCellData(Cell cell) {
		Object value = null;
		if (cell.getCellTypeEnum() == CellType.NUMERIC) {
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				SimpleDateFormat sdf = new SimpleDateFormat(Constants.DATE_FORMAT_DDMMYYYY);
				value = sdf.format(date);
			} else {
				value = Double.toString(cell.getNumericCellValue());
			}
		} else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
			value = Boolean.toString(cell.getBooleanCellValue());
		} else {
			value = cell.getStringCellValue();
		}
		return value;
	}

}
