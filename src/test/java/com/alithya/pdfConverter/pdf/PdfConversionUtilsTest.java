package com.alithya.pdfConverter.pdf;

import static org.junit.Assert.assertTrue;

import org.junit.Before;
import org.junit.Test;

import com.alithya.proof.poi.pdf.PdfConversionUtils;

public class PdfConversionUtilsTest {
	private static final String INPUT_FILE_PATH = "C:\\telus\\poi\\Leila.docx";
	private static final String OUTPUT_PDF_FILE_PATH = "C:\\telus\\poi\\GeneratedPDF.pdf";
	private static final String INPUT_EXCEL_FILE_PATH = "C:\\telus\\poi\\Sample.xlsx";
	private static final String OUTPUT_FILE_EXC_PDF_PATH = "C:\\telus\\poi\\GeneratedPDFfromExcel.pdf";

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public final void testConvertWordToPdf() {
		assertTrue(PdfConversionUtils.convertWordToPdf(INPUT_FILE_PATH, OUTPUT_PDF_FILE_PATH));

	}

	@Test
	public final void testReadExcelSheetAndWritePdf() {
		assertTrue(PdfConversionUtils.readExcelSheetAndWritePdf(INPUT_EXCEL_FILE_PATH, 0, OUTPUT_FILE_EXC_PDF_PATH));
	}

}
