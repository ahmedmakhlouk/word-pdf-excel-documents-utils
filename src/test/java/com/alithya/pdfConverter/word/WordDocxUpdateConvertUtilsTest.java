package com.alithya.pdfConverter.word;

import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;

import com.alithya.proof.poi.word.model.ParagraphsObject;
import com.alithya.proof.poi.word.process.WordDocxUpdateConvertUtils;

public class WordDocxUpdateConvertUtilsTest {

	private static final String INPUT_FILE_PATH = "C:\\telus\\poi\\Leila.docx";
	private static final String OUTPUT_FILE_PATH = "C:\\telus\\poi\\GeneratedWord.docx";
	private static final String OUTPUT_PDF_FILE_PATH = "C:\\telus\\poi\\GeneratedPDF.pdf";

	@Before
	public void setUp() throws Exception {

	}

	public static class MyData {
		private long providerNumber = 917; // "0000917"
		private String pharmacyName = "PHARMA_TEL";
		private String location = "Box 51, 3700 Anderson \n\t Calgary, AB T2W 3G4 \n";
		private String auditReason = "PSHCP Profiling";
		private String auditDate = "April 9, 2015";
		private String auditType = "Desk audit";
		private String auditor = "Van Pham";

		public long getProviderNumber() {
			return providerNumber;
		}

		public String getPharmacyName() {
			return pharmacyName;
		}

		public String getLocation() {
			return location;
		}

		public String getAuditReason() {
			return auditReason;
		}

		public String getAuditDate() {
			return auditDate;
		}

		public String getAuditType() {
			return auditType;
		}

		public String getAuditor() {
			return auditor;
		}
	}

	public static class MyObject {
		private String amount = "52,453.02";
		private String numberOfClaims = "50";
		private MyData testData = new MyData();

		public MyData getTestData() {
			return testData;
		}

		public String getAmount() {
			return amount;
		}

		public String getNumberOfClaims() {
			return numberOfClaims;
		}
	}

	@Test
	public final void testProcessReadAndUpdateWordFile() {
		try {
			WordDocxUpdateConvertUtils.processReadAndUpdateWordFile(INPUT_FILE_PATH, OUTPUT_FILE_PATH,
					getMockParagraphList(), getReplacementsMap());
		} catch (IOException e) {
			assertTrue(false);
		}
		assertTrue(true);
	}

	@Test
	public final void testProcessConversionWordToPdf() {
		assertTrue(WordDocxUpdateConvertUtils.processConversionWordToPdf(OUTPUT_FILE_PATH, OUTPUT_PDF_FILE_PATH));
	}

	/**
	 * mock map replacement data
	 * 
	 * @return
	 */
	public static Map<String, Object> getReplacementsMap() {
		Map<String, Object> replacementsMap = new HashMap<String, Object>() {
			private static final long serialVersionUID = 1L;
			{
				put("data", new MyData());
				put("object", new MyObject());
				put("auditRange", "1-6");
				put("satisfactionStatus",
						"Overall TELUS Health Solutions is not satisfied that this Pharmacy is conducting business within normal business guidelines");
				put("telusname", "Telus HEALTH");
				put("year01", "1554");
				put("year02", "1554");
				put("year03", "1554");
			}
		};
		return replacementsMap;
	}

	/**
	 * MOCK
	 * 
	 * @return
	 */
	private static List getMockParagraphList() {
		List listeOut = new ArrayList<ParagraphsObject>();
		////////////////////////////////////////////////////////////////////////////
		ParagraphsObject vElement = new ParagraphsObject();
		vElement.setTitle("No Evidence of Physician Authorization");
		vElement.setParagraphe(
				"Of the {year} claims audited, {prescriptionNbr} prescriptions processed at {phamacy} during the audited period did not have a corresponding recorded physician’s authorization, in that either there was no evidence of physician authorization and/or there were no physician authorized repeats.  For submission of all claims to TELUS Health, it is required that documentation be available on all verbal prescriptions and verbal authorizations for refills on both prescription requiring and OTC claims. Verbal prescriptions must be received from a licensed physician/dentist and must be reduced to writing by a pharmacist, prior to processing the claim.");
		listeOut.add(vElement);
		///////////////////////////////////////////////////////////
		vElement = new ParagraphsObject();
		vElement.setTitle("Incomplete Verbal Order Documentation");
		vElement.setParagraphe(
				"Of the {year} claims audited, {prescriptionNbr} prescriptions processed at {phamacy} during the audited period had incomplete documentation for a verbal authorization as per TELUS Health policies. TELUS Health requires that documentation be available on all verbal prescriptions and verbal repeats authorizations on both prescription requiring and OTC claims. Verbal prescriptions must be received from a licensed physician/dentist, and must be reduced to writing by a pharmacist. The documentation must include: The date the verbal prescription was received, the patient’s name, the drug quantity, directions for use, name of the prescriber, name of the receiving pharmacist, number of authorized repeats (if applicable) and the interval between refills (if applicable). Verbal authorizations for refills on both prescription requiring and OTC claims must be documented on the computer generated hardcopy, and must include the date the refill authorization was given and the number of refills authorized.  ");
		listeOut.add(vElement);
		/////////////////////////////////////////////////////////
		vElement = new ParagraphsObject();
		vElement.setTitle("Failure to Produce Documentation");
		vElement.setParagraphe(
				"Of the {year} claims audited, {prescriptionNbr} prescription processed at {phamacy} during the audited period, were charged back as Main Drug Mart failed to provide any corresponding documentation (physician’s authorization and/or corresponding computer generated hardcopy) to support the on-line processing of these claims. For submission of all claims to TELUS Health, it is required that documentation be available on all verbal prescriptions and verbal authorizations for refills on both prescription requiring and OTC claims. Verbal prescriptions must be received from a licensed physician/dentist and must be reduced to writing by a pharmacist, prior to processing the claim");
		listeOut.add(vElement);
		/////////////////////////////////////////////////////////
		vElement = new ParagraphsObject();
		vElement.setTitle("No Authorized Repeats");
		vElement.setParagraphe(
				"Of the {year} claims audited, {prescriptionNbr} prescriptions processed at {phamacy} during the audited period did not have any authorized repeats. For submission of all claims to TELUS Health, it is required that documentation be available on all verbal prescriptions and verbal authorizations for refills on both prescription requiring and OTC claims. Verbal prescriptions must be received from a licensed physician/dentist and must be reduced to writing by a pharmacist, prior to processing the claim");
		listeOut.add(vElement);
		/////////////////////////////////////////////////////////
		return listeOut;
	}

}
