package com.flex.assenttemplate.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.flex.assenttemplate.dto.BomTemplate;
import com.flex.assenttemplate.dto.MstrDetails;

public class ValidationUtil {

	// column index starts from 0
	private static final int MANUFACTURER_COLUMN_NUMBER = 2;
	private static final int MCODE_COLUMN_NUMBER = 3;
	private static final int EMAILID_COLUMN_NUMBER = 6;
	private static final int GLOBAL_MFR_COLUMN_NUMBER = 9;
	private static final int OBSOLETE_COLUMN_NUMBER = 10;

	public static void validateBomTemplate(String bomTemplateFileName, Integer bomTemplateFirstRow, String mstrFileName,
			Integer mstrFirstRow) throws IOException {

		System.out.println("..........................................................");
		System.out.println("Reading data from " + bomTemplateFileName + ". Please wait...");
		System.out.println("..........................................................");
		System.out.println("..........................................................");

		List<BomTemplate> bomTemplateData = getBomTemplateExcelData(bomTemplateFileName, bomTemplateFirstRow);

		System.out.println("Reading data from " + mstrFileName + ". Please wait...");
		System.out.println("..........................................................");
		System.out.println("..........................................................");
		System.out.println("..........................................................");
		Map<String, MstrDetails> mcodeAndMstrDetailsMap = getMstrExcelData(mstrFileName, mstrFirstRow);

		System.out.println("Validating mcode....");
		System.out.println("..........................................................");
		System.out.println("..........................................................");
		validateMcodeAndUpdateBomTemplate(bomTemplateFileName, bomTemplateData, mcodeAndMstrDetailsMap);
		System.out.println("Details are updated........................................");
		System.out.println("Please verify the file............" + bomTemplateFileName);
		System.out.println("..........................................................");

	}

	private static Map<String, MstrDetails> getMstrExcelData(String mstrFileName, Integer mstrFirstRow)
			throws IOException {

		// Read excel
		int lastColumn = 4;
		List<List<String>> rowList = FileUtil.readExcel(mstrFileName, mstrFirstRow, lastColumn);
		Map<String, MstrDetails> mcodeAndMstrDetailsMap = new HashMap<>();

		// Populate columns to java
		for (List<String> row : rowList) {

			MstrDetails mstrDetails = new MstrDetails();
			mstrDetails.setGlobalMfgCodes(row.get(0));
			mstrDetails.setGlobalManufacturerName(row.get(1));
			mstrDetails.setObsolete(row.get(3));

			mcodeAndMstrDetailsMap.put(mstrDetails.getGlobalMfgCodes(), mstrDetails);

		}
		return mcodeAndMstrDetailsMap;
	}

	private static List<BomTemplate> getBomTemplateExcelData(String bomTemplateFileName, Integer bomTemplateFirstRow)
			throws IOException {

		// Read excel
		int lastColumn = 7;
		List<List<String>> rowList = FileUtil.readExcel(bomTemplateFileName, bomTemplateFirstRow, lastColumn);
		List<BomTemplate> bomTemplateList = new ArrayList<>();

		// Populate columns to java
		for (List<String> row : rowList) {

			BomTemplate bomTemplate = new BomTemplate();
			bomTemplate.setFlexPartNo(row.get(0));
			bomTemplate.setDescription(row.get(1));
			bomTemplate.setManufacturer(row.get(2));
			bomTemplate.setMcode(row.get(3));
			bomTemplate.setMpn(row.get(4));
			bomTemplate.setEmailID(row.get(6));

			bomTemplateList.add(bomTemplate);

		}

		return bomTemplateList;
	}

	private static void validateMcodeAndUpdateBomTemplate(String bomTemplateFileName, List<BomTemplate> bomTemplateList,
			Map<String, MstrDetails> mcodeAndMstrDetailsMap) throws EncryptedDocumentException, IOException {

		FileInputStream inputStream = new FileInputStream(new File(bomTemplateFileName));
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(0);

		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// Get each row
		int rowNum = 1;
		for (BomTemplate bomTemplate : bomTemplateList) {

			if (bomTemplate.getEmailID().toLowerCase().contains("@flex.com")
					|| bomTemplate.getEmailID().toLowerCase().contains("@bomcheck.com")) {

				// Highlight email column in bom template excel
				String email = sheet.getRow(rowNum).getCell(EMAILID_COLUMN_NUMBER).getStringCellValue();

				Cell emailCell = sheet.getRow(rowNum).createCell(EMAILID_COLUMN_NUMBER);
				emailCell.setCellStyle(cellStyle);
				emailCell.setCellValue(email);

				// sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER).setCellStyle(null);

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

				if (!bomTemplate.getManufacturer().equals(mstrDetails.getGlobalManufacturerName())) {

					// highlight manufacturer cell in bom template excel

					String manufacturer = sheet.getRow(rowNum).getCell(MANUFACTURER_COLUMN_NUMBER).getStringCellValue();
					Cell manufacturerCell = sheet.getRow(rowNum).createCell(MANUFACTURER_COLUMN_NUMBER);
					manufacturerCell.setCellStyle(cellStyle);
					manufacturerCell.setCellValue(manufacturer);

					// sheet.getRow(rowNum).getCell(MANUFACTURER_COLUMN_NUMBER).setCellStyle(null);
				}

			} else {

				// Highlight mcode cell in bom template excel

				String mcode = sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER).getStringCellValue();
				Cell mcodeCell = sheet.getRow(rowNum).createCell(MCODE_COLUMN_NUMBER);
				mcodeCell.setCellStyle(cellStyle);
				mcodeCell.setCellValue(mcode);

				// sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER).setCellStyle(null);
			}

			rowNum++;

		}

		FileOutputStream outputStream = new FileOutputStream(bomTemplateFileName);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

	}

}
