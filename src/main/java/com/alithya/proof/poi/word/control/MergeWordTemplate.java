package com.alithya.proof.poi.word.control;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;

import com.alithya.proof.poi.word.model.ParagraphsObject;

public class MergeWordTemplate {

	private static final String REGEX_TARGET_TO_REPLACE = "\\$\\{(.*?)\\}";
	private static final String TARGET_BEGIN_DELIMETER = "${";
	private static final String TARGET_END_DELIMETER = "}";
	private static final String BLACK = "000000";
	private static final String ARIAL_FONT = "Arial";
	private static final String CALIBRE_FONT = "Calibre LIght";
	private static final String EMPTY = "";

	private XWPFDocument document;

	public MergeWordTemplate() {
	}

	/**
	 * @param dataMaps
	 * @param inputFilePath
	 * @param outputFilePAth
	 * @param listOfParagraphs
	 */
	public void modifyWordContentWithDataMap(Map<String, Object> dataMaps, String inputFilePath, String outputFilePAth,
			List listOfParagraphs) {
		// Read Word Content
		readWordFile(inputFilePath);

		// Modify Content
		mergeWithDataMap(buildDataMapToReplace(dataMaps), listOfParagraphs);

		// Update Word in the output file
		saveUpdatedDocument(outputFilePAth);
	}

	/**
	 * @param dataMaps
	 * @return
	 */
	private Map<String, String> buildDataMapToReplace(Map<String, Object> dataMaps) {
		Map<String, String> allInOnedataMap = new HashMap<String, String>();

		for (String myDataName : dataMaps.keySet()) {
			Object o = dataMaps.get(myDataName);
			if (o instanceof String) {
				allInOnedataMap.put(myDataName, (String) o);
			} else {
				drillDown(myDataName + ".", allInOnedataMap, o);
			}
		}
		return allInOnedataMap;
	}

	/**
	 * @param identifier
	 * @param allInOnedataMap
	 * @param o
	 */
	private void drillDown(String identifier, Map<String, String> allInOnedataMap, Object o) {
		Method[] getters = o.getClass().getDeclaredMethods();
		for (Method getter : getters) {
			String name = getter.getName();
			if (name.startsWith("get")) {
				String s4 = name.substring(3, 4).toLowerCase();
				String fieldName = s4 + name.substring(4);
				try {
					if (getter.getReturnType() == String.class || getter.getReturnType().isPrimitive()) {
						allInOnedataMap.put(identifier + fieldName, getter.invoke(o, (Object[]) null).toString());
					} else if (getter.getReturnType() == Date.class) {
						Date date = (Date) getter.invoke(o, (Object[]) null);
						String dateFormatted = "";
						allInOnedataMap.put(identifier + fieldName, dateFormatted);
					} else {
						drillDown(identifier + fieldName + ".", allInOnedataMap, getter.invoke(o, (Object[]) null));
					}
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * Load the word file into application document is the whole file
	 * 
	 * @param pathFile
	 */
	private void readWordFile(final String pathFile) {

		FileInputStream wordFile = null;
		try {

			wordFile = new FileInputStream(pathFile);
			document = new XWPFDocument(wordFile);
		}

		//
		catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				wordFile.close();
			} catch (IOException e) {
				/* TOO BAD */ }
		}
	}

	/**
	 * Save our updated file
	 * 
	 * @param filePath
	 */
	private void saveUpdatedDocument(final String filePath) {
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			document.write(out);
		}

		catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
				document = null;
			} catch (IOException e) {
				/* TOO BAD */ }
		}
	}

	/**
	 * Replace targets by their values indicated in the Map
	 * 
	 * @param listOfParagraphs
	 * 
	 * @param dataReplacement:
	 *            map<String target, String value>
	 */
	private void mergeWithDataMap(Map<String, String> dataReplacement, List listOfParagraphs) {
		for (XWPFParagraph paragraph : document.getParagraphs()) {
			if (isReplaceableText(paragraph.getText())) {
				List<XWPFRun> runs = paragraph.getRuns();
				if (runs != null) {
					for (XWPFRun r : runs) {
						String text = r.getText(0);
						if (isReplaceableText(text)) {
							String newText = replaceTargetsInText(text, dataReplacement);
							r.setText(newText, 0);
						}
					}
				}
			}
		}
		/***** create table of paragraphs given in the list ******/
		// createTableOfParagraphs(document, listOfParagraphs);
		createTableOfParagraphsNoTableOption(document, listOfParagraphs);
	}

	/**
	 * @param document2
	 * @param listOfParagraphs
	 */
	private void createTableOfParagraphs(XWPFDocument document2, List listOfParagraphs) {

		if (!listOfParagraphs.isEmpty()) {
			// Create new table in document
			XWPFTable newTemplateTable = document.createTable();
			newTemplateTable.removeRow(0);

			// Disable table borders
			newTemplateTable.getCTTbl().getTblPr().unsetTblBorders();
			/*********/
			CTTblGrid arg0;
			// arg0 = CTTblGrid.class.newInstance();
			arg0 = newTemplateTable.getCTTbl().addNewTblGrid();
			CTTblGridCol col0 = arg0.addNewGridCol();
			col0.setW(BigInteger.ONE);
			newTemplateTable.getCTTbl().setTblGrid(arg0);

			/*********/
			for (ParagraphsObject vParagraph : (List<ParagraphsObject>) listOfParagraphs) {
				// add new row
				XWPFTableRow newRow = newTemplateTable.createRow();
				if (newTemplateTable.getNumberOfRows() == 1) {
					newRow.addNewTableCell();
				}
				// Add Title
				newRow.getCell(0).setText(EMPTY); // initialize cell text
				setRun(newRow.getCell(0).addParagraph().createRun(), ARIAL_FONT, 12, BLACK, vParagraph.getTitle(), true,
						false);
				// Add text paragraph content
				newRow.getCell(0).setText(EMPTY); // initialize cell text
				setRun(newRow.getCell(0).addParagraph().createRun(), CALIBRE_FONT, 10, BLACK,
						vParagraph.getParagraphe(), false, true);
				newRow.getCell(0).removeParagraph(0);
			}
		}
	}

	/**
	 * @param document2
	 * @param listOfParagraphs
	 */
	private void createTableOfParagraphsNoTableOption(XWPFDocument document2, List listOfParagraphs) {

		if (!listOfParagraphs.isEmpty()) {

			for (ParagraphsObject vParagraph : (List<ParagraphsObject>) listOfParagraphs) {

				XWPFParagraph paragraphTitle = document2.createParagraph();
				setRun(paragraphTitle.createRun(), ARIAL_FONT, 12, BLACK, vParagraph.getTitle(), true, false);

				XWPFParagraph paragraphText = document2.createParagraph();
				setRun(paragraphText.createRun(), CALIBRE_FONT, 10, BLACK, vParagraph.getParagraphe(), false, true);
				document2.createParagraph();
			}
		}
	}

	/**
	 * @param run
	 * @param fontFamily
	 * @param fontSize
	 * @param colorRGB
	 * @param text
	 * @param bold
	 * @param addBreak
	 */
	private static void setRun(XWPFRun run, String fontFamily, int fontSize, String colorRGB, String text, boolean bold,
			boolean addBreak) {
		run.setFontFamily(fontFamily);
		run.setFontSize(fontSize);
		run.setColor(colorRGB);
		run.setText(text);
		run.setBold(bold);
		if (addBreak)
			run.addBreak();
	}

	/**
	 * @param document2
	 * @param dataReplacement
	 */
	private void readAndModifyDocumentTables(XWPFDocument document2, Map<String, String> dataReplacement) {

		/************ TABLES ************/
		List<XWPFTable> tables = document.getTables();

		System.out.println("Tables");
		for (XWPFTable table : tables) {

			for (XWPFTableRow row : table.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					String cellText = cell.getText();
					if (isReplaceableText(cellText)) {
						String newCellText = replaceTargetsInText(cellText, dataReplacement);
						cell.removeParagraph(0);
						cell.addParagraph();
						cell.setText(newCellText);

					}
				}
			}
		}
	}

	/**
	 * @param text
	 * @return
	 */
	private boolean isReplaceableText(String text) {
		return text != null && text.contains(TARGET_BEGIN_DELIMETER);
	}

	/**
	 * Replace parameter field with the given value
	 * 
	 * @param text
	 * @param mapVariablesValues
	 * @return
	 */
	public static String replaceTargetsInText(String text, Map<String, String> mapVariablesValues) {
		String newText = text;

		List<String> targetsList = getAllTargets(REGEX_TARGET_TO_REPLACE, text);
		for (String key : targetsList) {
			if (mapVariablesValues.keySet().contains(key)) {
				System.out.println("input  : " + text);
				System.out.print("\n");
				newText = newText.replace(TARGET_BEGIN_DELIMETER + key + TARGET_END_DELIMETER,
						mapVariablesValues.get(key));
				System.out.println("output : " + newText);
				System.out.print("\n");
			}
		}

		return newText;
	}

	/**
	 * Get targets to replace from text a target is a string inside ${.....}
	 * 
	 * @param text
	 * @return
	 */
	public static List<String> getAllTargets(String regexTarget, String text) {
		List<String> matches = new ArrayList<String>();
		Matcher m = Pattern.compile(regexTarget).matcher(text);
		while (m.find()) {
			matches.add(m.group(1));
		}
		return matches;
	}

}
