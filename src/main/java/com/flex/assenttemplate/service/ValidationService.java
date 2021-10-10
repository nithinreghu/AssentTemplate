package com.flex.assenttemplate.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.flex.assenttemplate.dto.BomTemplate;
import com.flex.assenttemplate.dto.MstrDetails;
import com.flex.assenttemplate.util.FileUtil;

@Service
public class ValidationService {

	@Value("${assentTemplateFileName}")
	private String assentTemplateFileName;

	@Value("${bomTemplateFileName}")
	private String bomTemplateFileName;

	@Value("${bomTemplateFirstRow:2}")
	private Integer bomTemplateFirstRow;

	@Value("${mstrFileName}")
	private String mstrFileName;

	@Value("${mstrFirstRow:2}")
	private Integer mstrFirstRow;

	// column index starts from 0
	private static final int MANUFACTURER_COLUMN_NUMBER = 2;
	private static final int MCODE_COLUMN_NUMBER = 3;
	private static final int EMAILID_COLUMN_NUMBER = 6;
	private static final int GLOBAL_MFR_COLUMN_NUMBER = 9;
	private static final int OBSOLETE_COLUMN_NUMBER = 10;
	private static final int USE_INSTEAD_COLUMN_NUMBER = 11;
	private static final int ALTERNATE_MFR_COLUMN_NUMBER = 12;

	private static final int REMARKS_COLUMN_NUMBER = 13;

	private static final String REGEX_PATTERN_FOR_EMAIL = "(^[ ]*[a-z0-9_.-]+@[a-z0-9.-]+[ ]*$)";

	public static final String VALIDATION_STATUS_SHEET_NAME = "Asset Validation QWERTY";

	public static final String SUCCESS = "Success";
	public static final String ERROR = "Error";

	private static final String YES = "yes";
	public static final String ERROR_FOUND = "Error Found";

	public void validateBomTemplate() throws IOException {

		System.out.println("..................................................................");
		System.out.println("........Reading data from " + bomTemplateFileName + ". Please wait...");
		System.out.println("..................................................................");
		System.out.println("..................................................................");

		int bomExcelLastColumnToBeRead = 10;
		List<BomTemplate> bomTemplateList = FileUtil.getBomTemplateExcelData(bomTemplateFileName,
				bomExcelLastColumnToBeRead);

		System.out.println("........Reading data from " + mstrFileName + ". Please wait...");
		System.out.println("..................................................................");
		System.out.println("..................................................................");
		System.out.println("..................................................................");
		Map<String, MstrDetails> mcodeAndMstrDetailsMap = getMstrExcelData(mstrFileName);

		System.out.println("........Validating mcode....");
		System.out.println("..................................................................");
		System.out.println("..................................................................");
		validateMcodeAndUpdateBomTemplate(bomTemplateFileName, bomTemplateList, mcodeAndMstrDetailsMap);
		System.out.println("........Details are updated........................................");
		System.out.println("........Please verify the file............" + bomTemplateFileName);
		System.out.println("..................................................................");

	}

	public Map<String, MstrDetails> getMstrExcelData(String mstrFileName) throws IOException {

		// Read excel
		int lastColumn = 7;
		List<List<String>> rowList = FileUtil.readExcel(mstrFileName, lastColumn);
		Map<String, MstrDetails> mcodeAndMstrDetailsMap = new HashMap<>();

		// Populate columns to java
		for (List<String> row : rowList) {

			MstrDetails mstrDetails = new MstrDetails();
			mstrDetails.setGlobalMfgCodes(row.get(0));
			mstrDetails.setGlobalManufacturerName(row.get(1));
			mstrDetails.setObsolete(row.get(3));
			mstrDetails.setUseInstead(row.get(6));

			mcodeAndMstrDetailsMap.put(mstrDetails.getGlobalMfgCodes(), mstrDetails);

		}
		return mcodeAndMstrDetailsMap;
	}

	private void validateMcodeAndUpdateBomTemplate(String bomTemplateFileName, List<BomTemplate> bomTemplateList,
			Map<String, MstrDetails> mcodeAndMstrDetailsMap) throws EncryptedDocumentException, IOException {

		FileInputStream inputStream = new FileInputStream(new File(bomTemplateFileName));
		Workbook workbook = WorkbookFactory.create(inputStream);

		Sheet sheet = workbook.getSheetAt(0);
		boolean errorFound = false;

		CellStyle cellStyle = updateCellStyle(workbook.createCellStyle());

		// Get each row
		int rowNum = 1;
		for (BomTemplate bomTemplate : bomTemplateList) {

			if (null == bomTemplate.getMpn() || bomTemplate.getMpn().isEmpty()) {
				System.out.println(
						"........BLANK MPN FOUND at row " + (rowNum + 1) + " in the File: " + bomTemplateFileName + "");
				System.out.println(
						"........IGNORING REMAINING rows. Please re-check the excel IF THIS IS NOT THE LAST ROW");
				System.out.println("..................................................................");
				System.out.println("..................................................................");
				break;
			}

			if (isInvalidEmail(bomTemplate.getEmailID().toLowerCase())) {

				errorFound = true;
				// Highlight email column in bom template excel
				String email = sheet.getRow(rowNum).getCell(EMAILID_COLUMN_NUMBER).getStringCellValue();

				Cell emailCell = sheet.getRow(rowNum).createCell(EMAILID_COLUMN_NUMBER);
				emailCell.setCellStyle(cellStyle);
				emailCell.setCellValue(email);

				sheet.getRow(rowNum).createCell(REMARKS_COLUMN_NUMBER).setCellValue(ERROR_FOUND);

			}

			// Get the values corresponding to mcode
			MstrDetails mstrDetails = mcodeAndMstrDetailsMap.get(bomTemplate.getMcode());
			if (null != mstrDetails) {

				// Update global mfr in bom template excel
				String mfrFromMstrExcel = mstrDetails.getGlobalManufacturerName();
				sheet.getRow(rowNum).createCell(GLOBAL_MFR_COLUMN_NUMBER).setCellValue(mfrFromMstrExcel);

				// Update obsolete in bom template excel
				String obsoleteFromMstrExcel = mstrDetails.getObsolete();
				sheet.getRow(rowNum).createCell(OBSOLETE_COLUMN_NUMBER).setCellValue(obsoleteFromMstrExcel);

				if (YES.equalsIgnoreCase(obsoleteFromMstrExcel)) {
					sheet.getRow(rowNum).createCell(USE_INSTEAD_COLUMN_NUMBER)
							.setCellValue(mstrDetails.getUseInstead());

					errorFound = true;

					if (null != mstrDetails.getUseInstead()) {

						// Get the values corresponding to alternate mcode
						MstrDetails alternateMstrDetails = mcodeAndMstrDetailsMap.get(mstrDetails.getUseInstead());
						if (null != alternateMstrDetails && null != alternateMstrDetails.getGlobalManufacturerName()) {
							// Update global mfr name in alternate mfr column
							sheet.getRow(rowNum).createCell(ALTERNATE_MFR_COLUMN_NUMBER)
									.setCellValue(alternateMstrDetails.getGlobalManufacturerName());
						}
					}

					sheet.getRow(rowNum).createCell(REMARKS_COLUMN_NUMBER).setCellValue(ERROR_FOUND);
				}

				if (!bomTemplate.getManufacturer().equals(mstrDetails.getGlobalManufacturerName())) {

					errorFound = true;
					// highlight manufacturer cell in bom template excel

					String manufacturer = sheet.getRow(rowNum).getCell(MANUFACTURER_COLUMN_NUMBER).getStringCellValue();
					Cell manufacturerCell = sheet.getRow(rowNum).createCell(MANUFACTURER_COLUMN_NUMBER);
					manufacturerCell.setCellStyle(cellStyle);
					manufacturerCell.setCellValue(manufacturer);

					sheet.getRow(rowNum).createCell(REMARKS_COLUMN_NUMBER).setCellValue(ERROR_FOUND);
				}

			} else {

				errorFound = true;
				// Highlight mcode cell in bom template excel

				String mcode = sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER).getStringCellValue();
				Cell mcodeCell = sheet.getRow(rowNum).createCell(MCODE_COLUMN_NUMBER);
				mcodeCell.setCellStyle(cellStyle);
				mcodeCell.setCellValue(mcode);

				sheet.getRow(rowNum).createCell(GLOBAL_MFR_COLUMN_NUMBER)
						.setCellValue("MCode details are not available in MSTR document");

				sheet.getRow(rowNum).createCell(REMARKS_COLUMN_NUMBER).setCellValue(ERROR_FOUND);

			}
			rowNum++;

		}

		setCompletionStatus(workbook, errorFound);

		FileOutputStream outputStream = new FileOutputStream(bomTemplateFileName);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

	}

	private void setCompletionStatus(Workbook workbook, boolean errorFound) {
		String sheetName = VALIDATION_STATUS_SHEET_NAME;
		Sheet sheet = workbook.getSheet(sheetName);

		int index = workbook.getSheetIndex(sheetName);
		if (index < 0) {
			// if sheet doesnt exist
			sheet = workbook.createSheet(sheetName);
		}

		Cell cell = sheet.createRow(0).createCell(0);
		if (errorFound) {

			cell.setCellValue(ERROR);
		} else {
			cell.setCellValue(SUCCESS);
		}

		workbook.setActiveSheet(0);
		workbook.setSheetHidden(workbook.getSheetIndex(sheetName), true);

		System.out.println("........Updated validation status....");
	}

	private CellStyle updateCellStyle(CellStyle cellStyle) {

		cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);

		return cellStyle;
	}

	private boolean isInvalidEmail(String emailID) {

		// Checking space
		if (emailID == null || emailID.isEmpty()) {
			return false;
		}

		boolean isInvalid = false;

		Pattern regexToFind = Pattern.compile(REGEX_PATTERN_FOR_EMAIL);

		Matcher regexMatcherToFind = regexToFind.matcher(emailID);
		if (!regexMatcherToFind.find()) {
			isInvalid = true;
		}

		return isInvalid || emailID.contains("@flex.com") || emailID.contains("@bomcheck.com");
	}

}
