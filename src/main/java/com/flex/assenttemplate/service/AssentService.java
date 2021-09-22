	package com.flex.assenttemplate.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.flex.assenttemplate.dto.BomTemplate;
import com.flex.assenttemplate.util.FileUtil;

@Service
public class AssentService {

	@Value("${assentTemplateFileName}")
	private String assentTemplateFileName;

	@Value("${bomTemplateFileName}")
	private String bomTemplateFileName;

	@Value("${customerName}")
	private String customerName;

	@Value("${ticketNumber}")
	private String ticketNumber;

	private static final int ASSENT_SUPPLIER_DETAILS_SUPPLIER_NAME = 3;
	private static final int ASSENT_SUPPLIER_DETAILS_SUPPLIER_NO = 4;
	private static final int ASSENT_SUPPLIER_DETAILS_TICKET_NUMBER = 21;

	private static final int ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NAME = 3;
	private static final int ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NO = 4;
	private static final int ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_CONTACT_EMAIL = 5;

	private static final int ASSENT_LEVELED_BOM_DETAILS_BOM_LEVEL = 3;
	private static final int ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NUMBER = 4;
	private static final int ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NAME = 5;
	private static final int ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_NAME = 7;
	private static final int ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_NUMBER = 8;
	private static final int ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_PART_NUMBER = 9;
	private static final int ASSENT_LEVELED_BOM_DETAILS_QUANTITY = 10;
	private static final int ASSENT_LEVELED_BOM_DETAILS_UNIT_OF_MEASURE = 11;
	private static final int ASSENT_LEVELED_BOM_DETAILS_STATUS = 21;
	private static final int ASSENT_LEVELED_BOM_DETAILS_FLEX_PART_NO = 27;
	private static final int ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME = 28;
	private static final int ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME = 29;
	private static final int ASSENT_LEVELED_BOM_DETAILS_TICKET_NO = 33;

	public void generateAssentTemplate() throws IOException {

		System.out.println("..................................................................");

		System.out.println("........Assent template file name..." + assentTemplateFileName);
		System.out.println("........Bom template file name......" + bomTemplateFileName);
		System.out.println("........Customer Name..............." + customerName);
		System.out.println("........Ticket Number..............." + ticketNumber);

		System.out.println("..................................................................");
		System.out.println("........Reading data from " + bomTemplateFileName + ". Please wait...");
		System.out.println("..................................................................");
		System.out.println("..................................................................");

		int bomExcelLastColumnToBeRead = 10;
		List<BomTemplate> bomTemplateList = FileUtil.getBomTemplateExcelData(bomTemplateFileName,
				bomExcelLastColumnToBeRead);

		updateAssentTemplate(bomTemplateList, assentTemplateFileName);

		System.out.println("........Template created successfully..............................");
		System.out.println("........Please refer the file............" + assentTemplateFileName);

	}

	private void updateAssentTemplate(List<BomTemplate> bomTemplateList, String assentTemplateFileName)
			throws EncryptedDocumentException, IOException {

		FileInputStream inputStream = new FileInputStream(new File(assentTemplateFileName));
		Workbook workbook = WorkbookFactory.create(inputStream);
		DataFormat format = workbook.createDataFormat();

		Sheet sheet2 = workbook.getSheetAt(1);
		Sheet sheet3 = workbook.getSheetAt(2);
		Sheet sheet4 = workbook.getSheetAt(3);

		// Get each row data from bom template excel
		// Get each row
		int rowNum = 1;

		updateLeveledBomDetailsFirstRow(sheet4, format, customerName, ticketNumber);

		for (BomTemplate bomTemplate : bomTemplateList) {

			// Update the values from bom template in the assent template excel

			// ---------------Sheet 2 (Supplier Details)-----------------
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_SUPPLIER_NAME, bomTemplate.getManufacturer(), format);
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_SUPPLIER_NO, bomTemplate.getMcode(), format);
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_TICKET_NUMBER, ticketNumber, format);

			// ---------------Sheet 3 (Supplier Contact Details)---------
			updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NAME, bomTemplate.getManufacturer(),
					format);
			updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NO, bomTemplate.getMcode(), format);
			updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_CONTACT_EMAIL,
					bomTemplate.getEmailID(), format);

			// ---------------Sheet 3 (Leveled BOM Details) From second row --------
			int leveledBomRowNumber = rowNum + 1;
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_BOM_LEVEL, bomTemplate.getAssemblyNo(),
					format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NUMBER,
					bomTemplate.getMpn(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NAME,
					bomTemplate.getDescription(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_NAME,
					bomTemplate.getManufacturer(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_NUMBER,
					bomTemplate.getMcode(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_PART_NUMBER,
					bomTemplate.getFlexPartNo(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_QUANTITY, bomTemplate.getQuantity(),
					format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_UNIT_OF_MEASURE, "each", format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_STATUS, "qualified", format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_FLEX_PART_NO,
					bomTemplate.getFlexPartNo(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME, customerName, format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME,
					ticketNumber + "_" + customerName, format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_TICKET_NO, ticketNumber, format);

			rowNum++;
		}

		// Save excel
		FileOutputStream outputStream = new FileOutputStream(assentTemplateFileName);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

	}

	private void updateLeveledBomDetailsFirstRow(Sheet sheet, DataFormat format, String customerName,
			String ticketNumber) {

		updateColumn(sheet, 1, ASSENT_LEVELED_BOM_DETAILS_BOM_LEVEL, "0", format);
		updateColumn(sheet, 1, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NUMBER, ticketNumber + "_" + customerName,
				format);
		updateColumn(sheet, 1, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NAME, ticketNumber + "_" + customerName, format);
		updateColumn(sheet, 1, ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME, customerName, format);
		updateColumn(sheet, 1, ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME, ticketNumber + "_" + customerName, format);
		updateColumn(sheet, 1, ASSENT_LEVELED_BOM_DETAILS_TICKET_NO, ticketNumber, format);

	}

	/**
	 * Update the value in each cell.
	 * 
	 * @param sheet
	 * @param rowNum       row number starts from 0
	 * @param columnNumber number corresponding to colummn (eg: A -> 0, B->1...)
	 * @param columnValue  value to be updated in the column
	 * @param format
	 */
	private void updateColumn(Sheet sheet, int rowNum, int columnNumber, String columnValue, DataFormat format) {

		try {

			Row row = sheet.getRow(rowNum);
			if (null == row) {
				row = sheet.createRow(rowNum);
			}

			Cell cell = row.getCell(columnNumber);
			CellStyle style = null;

			if (null == cell) {
				sheet.getRow(rowNum).createCell(columnNumber).setCellValue(columnValue);
			} else {
				sheet.getRow(rowNum).getCell(columnNumber).setCellValue(columnValue);
			}

			style = sheet.getRow(rowNum).getCell(columnNumber).getCellStyle();
			style.setDataFormat(format.getFormat("@"));
			sheet.getRow(rowNum).getCell(columnNumber).setCellStyle(style);

		} catch (Exception e) {
			throw e;
		}

	}

}
