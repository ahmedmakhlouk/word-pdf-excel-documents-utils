package com.alithya.proof.poi.word.process;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import com.alithya.proof.poi.pdf.PdfConversionUtils;
import com.alithya.proof.poi.word.control.DocxModifyer;
import com.alithya.proof.poi.word.control.MergeWordTemplate;

public class WordDocxUpdateConvertUtils {
	private static final String INTERMEDIATE_FILE_PATH = "C:\\poi\\IntermediateWord.docx";
	private static Logger LOGGER = Logger.getLogger(WordDocxUpdateConvertUtils.class.getName());
	private static MergeWordTemplate mergeWordTemplate = new MergeWordTemplate();

	/**
	 * Convert a given word file to pdf format
	 * 
	 * @param filePath
	 */
	public static boolean processConversionWordToPdf(final String inputFilePath, final String outPutFile) {
		LOGGER.info("Begin processConversionWordToPdf");
		boolean result = PdfConversionUtils.convertWordToPdf(inputFilePath, outPutFile);
		LOGGER.info("Convertion done !!");
		LOGGER.info("End processConversionWordToPdf");
		return result;
	}

	/**
	 * Read a word file and modifie the variables in it using the given list
	 * 
	 * @param inputFilePath
	 * @param outputFilePath
	 * @param listOfParagraphs
	 * @param replacementsMap
	 * @throws IOException
	 */
	public static void processReadAndUpdateWordFile(final String inputFilePath, final String outputFilePath,
			final List listOfParagraphs, Map<String, Object> replacementsMap) throws IOException {
		LOGGER.info("Begin processReadAndUpdateWordFile");
		DocxModifyer.modify(inputFilePath, INTERMEDIATE_FILE_PATH);

		LOGGER.info("Update the word file");
		// Modify Word Content
		mergeWordTemplate.modifyWordContentWithDataMap(replacementsMap, INTERMEDIATE_FILE_PATH, outputFilePath,
				listOfParagraphs);
		LOGGER.info("END processReadAndUpdateWordFile");
	}

}