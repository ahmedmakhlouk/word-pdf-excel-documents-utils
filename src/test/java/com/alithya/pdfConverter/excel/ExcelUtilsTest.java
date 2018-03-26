/**
 * 
 */
package com.alithya.pdfConverter.excel;

import static org.junit.Assert.assertTrue;

import org.junit.Before;
import org.junit.Test;

import com.alithya.proof.poi.excel.ExcelUtils;

/**
 * @author x201927
 *
 */
public class ExcelUtilsTest {

	private static final Integer MOCK_NUMBER = 99999;
	private static final int SHEET_NUMBER = 0;
	private static final String NEW_FILE_PATH = "";
	private static final String PATH_FILE = "";
	private static final String MOCK_STRING = "NEW_VALUE_IN";

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
	}

	/**
	 * Test method for {@link com.alithya.proof.poi.excel.ExcelUtils#readAndUpdateExcelFile(java.lang.String, java.lang.Integer, java.lang.String, int, java.lang.String)}.
	 */
	@Test
	public final void testReadAndUpdateExcelFile() {
		assertTrue(ExcelUtils.readAndUpdateExcelFile(PATH_FILE, MOCK_NUMBER, MOCK_STRING, SHEET_NUMBER, NEW_FILE_PATH));
	}

}
