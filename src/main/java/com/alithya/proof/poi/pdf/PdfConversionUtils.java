package com.alithya.proof.poi.pdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

public class PdfConversionUtils {
	private static Logger LOGGER = Logger.getLogger(PdfConversionUtils.class.getName());

	/**
	 * Convert a given word file into Pdf file
	 * 
	 * @param inputFilePath
	 * @param outPutFile
	 */
	public static boolean convertWordToPdf(final String inputFilePath, final String outPutFile) {
		LOGGER.info("Begin convertWordToPdf");
		String filePath = inputFilePath;
		FileInputStream fInputStream;
		try {
			fInputStream = new FileInputStream(new File(filePath));
			XWPFDocument documentToConvert = new XWPFDocument(fInputStream);
			File outFile = new File(outPutFile);
			outFile.getParentFile().mkdirs();
			OutputStream outerStream = new FileOutputStream(outFile);
			PdfOptions optionsForPdf = PdfOptions.create().fontEncoding("windows-1250");
			PdfConverter.getInstance().convert(documentToConvert, outerStream, optionsForPdf);
		} catch (FileNotFoundException e) {
			LOGGER.severe("Fail : " + e.getMessage());
			e.printStackTrace();
			return false;
		} catch (IOException e) {
			LOGGER.severe("Fail : " + e.getMessage());
			e.printStackTrace();
			return false;
		}
		LOGGER.info("End convertWordToPdf");
		return true;

	}

	/**
	 * @param pathFile
	 * @param sheetNumber
	 * @param outpuFilePath
	 */
	public static boolean readExcelSheetAndWritePdf(String pathFile, int sheetNumber, String outpuFilePath) {
		LOGGER.info("Begin readExcelSheetAndWritePdf ");
		try {
			FileInputStream excelFile = new FileInputStream(new File(pathFile));
			Sheet datatypeSheet = new XSSFWorkbook(excelFile).getSheetAt(sheetNumber);

			Iterator<Row> rowIterator = datatypeSheet.iterator();
			Row currentRow = null;
			Iterator<Cell> cellIterator = null;

			Document document = new Document();
			OutputStream fileOut = new FileOutputStream(new File(outpuFilePath));
			PdfWriter writer = PdfWriter.getInstance(document, fileOut);
			document.open();
			writer.setPageEmpty(true);
			document.newPage();
			writer.setPageEmpty(true);

			Integer numberOfColumns = 0;
			cellIterator = datatypeSheet.getRow(0).iterator();
			// loop on cells
			while (cellIterator.hasNext()) {
				cellIterator.next();
				numberOfColumns++;
			}
			if (numberOfColumns > 0) {
				drawPdfTable(numberOfColumns, currentRow, rowIterator, cellIterator, document);
			}

			document.close();
			excelFile.close();

		} catch (Exception e) {
			LOGGER.severe("FAIL : " + e.getMessage());
			return false;
		}
		LOGGER.info("End successfully readExcelSheetAndWritePdf ");
		return true;
	}

	/**
	 * Construct table in PDF file
	 * 
	 * @param numberOfColumns
	 * @param currentRow
	 * @param iterator
	 * @param cellIterator
	 * @param document
	 * @throws DocumentException
	 */
	private static void drawPdfTable(Integer numberOfColumns, Row currentRow, Iterator<Row> iterator,
			Iterator<Cell> cellIterator, Document document) throws DocumentException {

		PdfPTable table = new PdfPTable(numberOfColumns);
		// Loop on Rows
		do {
			currentRow = iterator.next();
			cellIterator = currentRow.iterator();
			// loop on cells
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					table.addCell(String.valueOf(currentCell.getNumericCellValue()));
				} else {
					table.addCell(currentCell.getStringCellValue());
				}
			}
		} while (iterator.hasNext());

		document.add(table);
	}

}
