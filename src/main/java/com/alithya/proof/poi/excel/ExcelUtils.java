package com.alithya.proof.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	private static Logger LOGGER = Logger.getLogger(ExcelUtils.class.getName());

	/**
	 * @param pathFile
	 * @param newValueNum
	 * @param newValueString
	 * @param sheetNumber
	 * @param pathNewFile
	 */
	public static boolean readAndUpdateExcelFile(final String pathFile, final Integer newValueNum,
			final String newValueString, final int sheetNumber, final String pathNewFile) {
		LOGGER.info("Begin readAndUpdateExcelFile");
		try {

			FileInputStream excelFile = new FileInputStream(new File(pathFile));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(sheetNumber);
			Iterator<Row> iterator = datatypeSheet.iterator();
			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();

			while (iterator.hasNext()) {
				currentRow = iterator.next();
				cellIterator = currentRow.iterator();
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						currentCell.setCellValue(newValueNum);
					} else {
						currentCell.setCellValue(newValueString);
					}
				}
			}
			workbook.setForceFormulaRecalculation(true);
			String val = "" + Math.random();
			FileOutputStream outStream = new FileOutputStream(new File(pathNewFile + val.replace('.', '_') + ".xlsx"));
			workbook.write(outStream);
		} catch (FileNotFoundException e) {
			LOGGER.info("End eith error readAndUpdateExcelFile " + e.getMessage());
			e.printStackTrace();
			return false;
		} catch (IOException e) {
			LOGGER.info("End eith error readAndUpdateExcelFile " + e.getMessage());
			e.printStackTrace();
			return false;
		}
		LOGGER.info("End readAndUpdateExcelFile");
		return true;

	}

}