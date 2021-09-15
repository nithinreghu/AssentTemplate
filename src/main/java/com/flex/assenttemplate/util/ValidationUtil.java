package com.flex.assenttemplate.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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

		List<BomTemplate> bomTemplateData = getBomTemplateExcelData(bomTemplateFileName, bomTemplateFirstRow);
		Map<String, MstrDetails> mcodeAndMstrDetailsMap = getMstrExcelData(mstrFileName, mstrFirstRow);

		validateMcodeAndUpdateBomTemplate(bomTemplateFileName, bomTemplateData, mcodeAndMstrDetailsMap);

	}

	private static Map<String, MstrDetails> getMstrExcelData(String mstrFileName, Integer mstrFirstRow)
			throws IOException {

		// Read excel
		List<List<String>> rowList = FileUtil.readExcel(mstrFileName, mstrFirstRow);
		Map<String, MstrDetails> mcodeAndMstrDetailsMap = new HashMap<>();

		// Populate columns to java
		for (List<String> row : rowList) {

			MstrDetails mstrDetails = new MstrDetails();
			mstrDetails.setGlobalMfgCodes(row.get(0));
			mstrDetails.setGlobalManufacturerName(row.get(1));
			mstrDetails.setObsolete(row.get(2));

			mcodeAndMstrDetailsMap.put(mstrDetails.getGlobalMfgCodes(), mstrDetails);

		}
		return mcodeAndMstrDetailsMap;
	}

	private static List<BomTemplate> getBomTemplateExcelData(String bomTemplateFileName, Integer bomTemplateFirstRow)
			throws IOException {

		// Read excel
		List<List<String>> rowList = FileUtil.readExcel(bomTemplateFileName, bomTemplateFirstRow);
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

		}

		return bomTemplateList;
	}

	private static void validateMcodeAndUpdateBomTemplate(String bomTemplateFileName, List<BomTemplate> bomTemplateList,
			Map<String, MstrDetails> mcodeAndMstrDetailsMap) throws EncryptedDocumentException, IOException {

		FileInputStream inputStream = new FileInputStream(new File(bomTemplateFileName));
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(0);

		// Get each row
		int rowNum = 1;
		for (BomTemplate bomTemplate : bomTemplateList) {

			if (bomTemplate.getEmailID().toLowerCase().contains("@flex.com")
					|| bomTemplate.getEmailID().toLowerCase().contains("@bomcheck.com")) {

				// Highlight email column in bom template excel
				Cell emailCell = sheet.getRow(rowNum).getCell(EMAILID_COLUMN_NUMBER);
				
				CellStyle emailCellStyle = emailCell.getCellStyle();
				emailCellStyle.setFillBackgroundColor(IndexedColors.YELLOW.index);
				
				emailCell.setCellStyle(emailCellStyle);
				// sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER).setCellStyle(null);

			}

			// Get the values corresponding to mcode
			MstrDetails mstrDetails = mcodeAndMstrDetailsMap.get(bomTemplate.getMcode());
			if (null != mstrDetails) {

				// Update global mfr in bom template excel
				String mfrFromMstrExcel = mstrDetails.getGlobalManufacturerName();
				sheet.getRow(rowNum).getCell(GLOBAL_MFR_COLUMN_NUMBER).setCellValue(mfrFromMstrExcel);

				// Update obsolete in bom template excel
				String obsoleteFromMstrExcel = mstrDetails.getObsolete();
				sheet.getRow(rowNum).getCell(OBSOLETE_COLUMN_NUMBER).setCellValue(obsoleteFromMstrExcel);

				if (!bomTemplate.getManufacturer().equals(mstrDetails.getGlobalManufacturerName())) {

					// highlight manufacturer cell in bom template excel

					Cell manufacturerCell = sheet.getRow(rowNum).getCell(MANUFACTURER_COLUMN_NUMBER);
					CellStyle manufacturerCellCellStyle = manufacturerCell.getCellStyle();
					manufacturerCellCellStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
					manufacturerCell.setCellStyle(manufacturerCellCellStyle);

					// sheet.getRow(rowNum).getCell(MANUFACTURER_COLUMN_NUMBER).setCellStyle(null);
				}

			} else {

				// Highlight mcode cell in bom template excel

				Cell mcodeCell = sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER);
				CellStyle mcodeCellStyle = mcodeCell.getCellStyle();
				mcodeCellStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
				mcodeCell.setCellStyle(mcodeCellStyle);

				// sheet.getRow(rowNum).getCell(MCODE_COLUMN_NUMBER).setCellStyle(null);
			}

			rowNum++;

		}

	}

}
