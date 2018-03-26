package com.alithya.proof.poi.word.control;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class DocxModifyer {
	private static final String FILE_TO_MODIFY = "word/document.xml";

	private static final String EMPTY = "";
	private static final String XML_BEGIN_TAG = "<";
	private static final String XML_END_TAG = ">";
	private static final String REGEX_XML_TAG_TO_REMOVE = "\\<(.*?)\\>";
	// all inside ${ ...... }
	private static final String REGEX_TARGET_TO_REPLACE = "\\$\\{(.*?)\\}";
	private static final String TARGET_BEGIN_DELIMETER = "${";
	private static final String TARGET_END_DELIMETER = "}";
	
	public static void modify(String inputPath, String outputPath) throws IOException {

		FileOutputStream fos = new FileOutputStream(outputPath);
		final ZipOutputStream zos = new ZipOutputStream(fos);

		ZipFile zipFile = new ZipFile(inputPath);
		for(Enumeration<? extends ZipEntry> e = zipFile.entries(); e.hasMoreElements(); ) {
//		    ZipEntry entryIn = (ZipEntry) e.nextElement();
		    ZipEntry entryIn = new ZipEntry((  (ZipEntry)e.nextElement()).getName());
	        zos.putNextEntry(entryIn);
	        InputStream is = zipFile.getInputStream(entryIn);

		    // move other files to new ZIP
		    if (  !entryIn.getName().equalsIgnoreCase(FILE_TO_MODIFY)  ) {   // (true) { // 
		    	justCopy(is , zos);
		    }

		    // Modify then move to new ZIP
		    else {
		    	modifyAndCopy(is , zos);
		    }

		    zos.closeEntry();
		}

		fos.close();
//		zos.close();
	}

	private static void justCopy(InputStream is , final ZipOutputStream zos) throws IOException {
        byte[] buf = new byte[1024];
		int len;
        while((len = is.read(buf)) > 0) {            
            zos.write(buf, 0, len);
        }
	}
	
	private static void modifyAndCopy(InputStream is , final ZipOutputStream zos) throws IOException {
        byte[] buf = new byte[1024];
		int length;

		StringBuilder textBuilder =  new StringBuilder();
		int bigLength = 0;

        while ((length = (is.read(buf))) > 0) {
    		textBuilder.append( new String(buf, 0, length) ); 
        	bigLength += length;
        }

        String s1 = textBuilder.toString();
        // get polluted targets
        List<String>  to_be_cleaned_list1 = getAllTargets(REGEX_TARGET_TO_REPLACE, s1);
        if ( !to_be_cleaned_list1.isEmpty() ) {
            // Build clean targets
            Map<String, String> cleaningMap = new HashMap<String, String>();
            for (String to_be_cleand : to_be_cleaned_list1) {
            	cleaningMap.put(TARGET_BEGIN_DELIMETER + to_be_cleand + TARGET_END_DELIMETER, 
            			TARGET_BEGIN_DELIMETER + cleanText(to_be_cleand) + TARGET_END_DELIMETER);
            }

            // Replace polluted targets by cleaned ones
            for (String key : cleaningMap.keySet()) {
            	s1 = s1.replace(key, cleaningMap.get(key));
            }
        }
        
        byte[] bigBuffer1 = new byte[bigLength];
        bigBuffer1 = s1.getBytes();
        zos.write(bigBuffer1, 0, bigBuffer1.length); 
	}
	

	/**
	 * Replace parameter field with the given value
	 * 
	 * @param text
	 * @param mapVariablesValues
	 * @return
	 */
	public static String cleanText(String text) {
		return text.replaceAll(REGEX_XML_TAG_TO_REMOVE, EMPTY).replaceAll(XML_BEGIN_TAG, EMPTY).replaceAll(XML_END_TAG, EMPTY);
	}

	
	/**
	 * Get targets to replace from text
	 * a target is a string inside ${.....}
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
	public static List<String> getAllTargets( String text) {
		return getAllTargets(REGEX_TARGET_TO_REPLACE, text);
	}
}
