package com.swapson.TestCaseReader;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.swapson.Util.ExcelReader;

/**
 * Test run
 * @author Swapnil Sonar
 *
 */
public class Main {
	private static final String FILE_PATH = "./src/main/resources/TestCaseData1.xlsx";

	public static void main(String[] args) {
		ExcelReader er;
		try {
			er = new ExcelReader(FILE_PATH);
			Map<String,List<Map<String,String>>> allTestData = er.readAllTestData(false);
			System.out.println("=== Reference Data ===");
			System.out.println(ExcelReader.REFERENCE_DATA);
			
			System.out.println("=== Test Data ===");
			System.out.println(allTestData);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}
}
