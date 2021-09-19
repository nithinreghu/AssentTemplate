package com.flex.assenttemplate.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
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
		System.out.println("..........................................................");
		System.out.println("Reading data from " + bomTemplateFileName + ". Please wait...");
		System.out.println("..........................................................");
		System.out.println("..........................................................");

		int bomExcelLastColumnToBeRead = 10;
		List<BomTemplate> bomTemplateList = FileUtil.getBomTemplateExcelData(bomTemplateFileName,
				bomExcelLastColumnToBeRead);

		updateAssentTemplate(bomTemplateList, assentTemplateFileName);

	}

	private void updateAssentTemplate(List<BomTemplate> bomTemplateList, String assentTemplateFileName)
			throws EncryptedDocumentException, IOException {

		FileInputStream inputStream = new FileInputStream(new File(assentTemplateFileName));
		Workbook workbook = WorkbookFactory.create(inputStream);

		Sheet sheet2 = workbook.getSheetAt(1);
		Sheet sheet3 = workbook.getSheetAt(2);
		Sheet sheet4 = workbook.getSheetAt(3);

		// Get each row data from bom template excel
		// Get each row
		int rowNum = 1;
		for (BomTemplate bomTemplate : bomTemplateList) {

			// Update the values from bom template in the assent template excel

			// ---------------Sheet 2 (Supplier Details)-----------------
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_SUPPLIER_NAME, bomTemplate.getManufacturer());
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_SUPPLIER_NO, bomTemplate.getMcode());

			// ---------------Sheet 3 (Supplier Contact Details)---------
			updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NAME, bomTemplate.getManufacturer());
			updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NO, bomTemplate.getMcode());
			updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_CONTACT_EMAIL,
					bomTemplate.getEmailID());

			// ---------------Sheet 3 (Leveled BOM Details) first row --------

			if (rowNum == 1) {
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_BOM_LEVEL, "0");
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NUMBER,
						ticketNumber + "_" + customerName);
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NAME,
						ticketNumber + "_" + customerName);
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME, customerName);
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME,
						ticketNumber + "_" + customerName);
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_TICKET_NO, ticketNumber);

			} else {
				// ---------------Sheet 3 (Leveled BOM Details) From second row --------

				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_BOM_LEVEL, bomTemplate.getAssemblyNo());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NUMBER, bomTemplate.getMpn());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_PRODUCT_PART_NAME,
						bomTemplate.getDescription());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_NAME, bomTemplate.getManufacturer());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_NUMBER, bomTemplate.getMcode());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_SUPPLIER_PART_NUMBER,
						bomTemplate.getFlexPartNo());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_QUANTITY, bomTemplate.getQuantity());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_UNIT_OF_MEASURE, "each");
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_STATUS, "qualified");
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_FLEX_PART_NO, bomTemplate.getFlexPartNo());
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME, customerName);
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME,
						ticketNumber + "_" + customerName);
				updateColumn(sheet4, rowNum, ASSENT_LEVELED_BOM_DETAILS_TICKET_NO, ticketNumber);
			}

			rowNum++;
		}

		// Save excel
		FileOutputStream outputStream = new FileOutputStream(bomTemplateFileName);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

	}

	/**
	 * Update the value in each cell.
	 * 
	 * @param sheet
	 * @param rowNum       row number starts from 0
	 * @param columnNumber number corresponding to colummn (eg: A -> 0, B->1...)
	 * @param columnValue  value to be updated in the column
	 */
	private void updateColumn(Sheet sheet, int rowNum, int columnNumber, String columnValue) {

		Cell cell = sheet.getRow(rowNum).getCell(columnNumber);
		if (null == cell) {

			sheet.getRow(rowNum).createCell(columnNumber).setCellValue(columnValue);
		} else {
			sheet.getRow(rowNum).getCell(columnNumber).setCellValue(columnValue);
		}

	}

}
