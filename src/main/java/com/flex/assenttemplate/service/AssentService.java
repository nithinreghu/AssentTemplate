package com.flex.assenttemplate.service;

import static com.flex.assenttemplate.util.Constants.OBSOLETE_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.SUCCESS;
import static com.flex.assenttemplate.util.Constants.VALIDATION_STATUS_SHEET_NAME;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.flex.assenttemplate.dto.BomTemplate;
import com.flex.assenttemplate.dto.SupplierContactDetails;
import com.flex.assenttemplate.dto.SupplierDetails;
import com.flex.assenttemplate.exception.MyException;
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
	private static final int ASSENT_LEVELED_BOM_DETAILS_COMMODITY = 25;
	private static final int ASSENT_LEVELED_BOM_DETAILS_STATUS = 26;
	private static final int ASSENT_LEVELED_BOM_DETAILS_FLEX_PART_NO = 32;
	private static final int ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME = 33;
	private static final int ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME = 34;
	private static final int ASSENT_LEVELED_BOM_DETAILS_TICKET_NO = 38;

	public void generateAssentTemplate() throws IOException, EncryptedDocumentException, MyException {

		System.out.println("..................................................................");

		System.out.println("........Assent template file name..." + assentTemplateFileName);
		System.out.println("........Bom template file name......" + bomTemplateFileName);
		System.out.println("........Customer Name..............." + customerName);
		System.out.println("........Ticket Number..............." + ticketNumber);

		System.out.println("..................................................................");
		System.out.println("........Verifying validation status in " + bomTemplateFileName + ". Please wait...");
		verifyValidationStatus();

		System.out.println("..................................................................");
		System.out.println("........Reading data from " + bomTemplateFileName + ". Please wait...");
		System.out.println("..................................................................");
		System.out.println("..................................................................");

		int bomExcelLastColumnToBeRead = OBSOLETE_COLUMN_NUMBER;
		List<BomTemplate> bomTemplateList = FileUtil.getBomTemplateExcelData(bomTemplateFileName,
				bomExcelLastColumnToBeRead);

		updateAssentTemplate(bomTemplateList, assentTemplateFileName);

		System.out.println("........Template created successfully..............................");
		System.out.println("........Please refer the file............" + assentTemplateFileName);

	}

	private void verifyValidationStatus() throws EncryptedDocumentException, IOException, MyException {

		FileInputStream inputStream = new FileInputStream(new File(bomTemplateFileName));
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheet(VALIDATION_STATUS_SHEET_NAME);

		if (null != sheet && null != sheet.getRow(0) && null != sheet.getRow(0).getCell(0)
				&& SUCCESS.equals(sheet.getRow(0).getCell(0).getStringCellValue())) {

			System.out.println("........Verified validation status... Proceeding to generate asset template");
		} else {
			System.out.println("..................................................................");
			System.out.println("..................................................................");
			System.out.println(
					"........Please verify if the BOM template valdation has been completed and all errors have been resolved.......");
			System.out.println(
					".........Make sure to RE-RUN the BOM template validation after resolving the errors (if there are any).......");
			System.out.println("..................................................................");
			System.out.println("........KINDLY VALIDATE BOM TEMPLATE AGAIN !!!!!!!!...............");

			throw new MyException();
		}

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

			// ---------------Sheet 4 (Leveled BOM Details) From second row --------
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
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_COMMODITY, bomTemplate.getCommodity(),
					format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_STATUS, "qualified", format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_FLEX_PART_NO,
					bomTemplate.getFlexPartNo(), format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_CUSTOMER_NAME, customerName, format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_PROJECT_NAME,
					ticketNumber + "_" + customerName, format);
			updateColumn(sheet4, leveledBomRowNumber, ASSENT_LEVELED_BOM_DETAILS_TICKET_NO, ticketNumber, format);

			rowNum++;
		}

		updateSupplierDetails(sheet2, format, bomTemplateList);
		updateSupplierContactDetails(sheet3, format, bomTemplateList);

		// Save excel
		FileOutputStream outputStream = new FileOutputStream(assentTemplateFileName);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

	}

	private void updateSupplierContactDetails(Sheet sheet3, DataFormat format, List<BomTemplate> bomTemplateList) {

		ModelMapper modelMapper = new ModelMapper();
		Set<SupplierContactDetails> uniqueSupplierContactDetails = bomTemplateList.stream()
				.map(bomTemplate -> modelMapper.map(bomTemplate, SupplierContactDetails.class))
				.collect(Collectors.toSet());

		int rowNum = 1;

		for (SupplierContactDetails supplierContactDetails : uniqueSupplierContactDetails) {

			// Update the values from bom template in the assent template excel

			// ---------------Sheet 3 (Supplier Contact Details)---------

			if (null != supplierContactDetails.getEmailID() && !supplierContactDetails.getEmailID().isEmpty()) {

				updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NAME,
						supplierContactDetails.getManufacturer(), format);
				updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_NO,
						supplierContactDetails.getMcode(), format);
				updateColumn(sheet3, rowNum, ASSENT_SUPPLIER_CONTACT_DETAILS_SUPPLIER_CONTACT_EMAIL,
						supplierContactDetails.getEmailID(), format);

				rowNum++;
			}
		}

	}

	private void updateSupplierDetails(Sheet sheet2, DataFormat format, List<BomTemplate> bomTemplateList) {

		ModelMapper modelMapper = new ModelMapper();
		Set<SupplierDetails> uniqueSupplierDetails = bomTemplateList.stream()
				.map(bomTemplate -> modelMapper.map(bomTemplate, SupplierDetails.class)).collect(Collectors.toSet());

		int rowNum = 1;

		for (SupplierDetails supplierDetails : uniqueSupplierDetails) {

			// Update the values from bom template in the assent template excel

			// ---------------Sheet 2 (Supplier Details)-----------------
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_SUPPLIER_NAME, supplierDetails.getManufacturer(),
					format);
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_SUPPLIER_NO, supplierDetails.getMcode(), format);
			updateColumn(sheet2, rowNum, ASSENT_SUPPLIER_DETAILS_TICKET_NUMBER, ticketNumber, format);
			rowNum++;
		}

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
	 * @param rowNum
	 *            row number starts from 0
	 * @param columnNumber
	 *            number corresponding to colummn (eg: A -> 0, B->1...)
	 * @param columnValue
	 *            value to be updated in the column
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
