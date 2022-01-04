package com.flex.assenttemplate.util;

import static com.flex.assenttemplate.util.Constants.ASSEMBLY_NO_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.COMMODITY_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.DESCRIPTION_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.EMAILID_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.FLEX_PART_NO_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.MANUFACTURER_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.MCODE_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.MPN_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.QUANTITY_COLUMN_NUMBER;
import static com.flex.assenttemplate.util.Constants.TEL_NO_COLUMN_NUMBER;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.flex.assenttemplate.dto.BomTemplate;

public class FileUtil {

	public static List<List<String>> readExcel(String filename, int lastColumn) throws IOException {

		List<List<String>> rowList = new ArrayList<>();
		Workbook workbook = null;
		try {
			List<String> columnList = null;
			workbook = WorkbookFactory.create(new File(filename));
			Sheet sheet = workbook.getSheetAt(0);

			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			for (int i = 1; i <= sheet.getLastRowNum(); i++) {

				columnList = new ArrayList<>();
				Row row = sheet.getRow(i);

				for (int j = 0; j < lastColumn; j++) {
					String cellValue = dataFormatter.formatCellValue(row.getCell(j));
					columnList.add(cellValue);
				}
				rowList.add(columnList);
			}

		} catch (IOException e) {
			throw e;
		} finally {
			if (null != workbook) {
				workbook.close();
			}
		}
		return rowList;

	}

	public static List<BomTemplate> getBomTemplateExcelData(String bomTemplateFileName, int lastColumn)
			throws IOException {

		// Read excel
		List<List<String>> rowList = readExcel(bomTemplateFileName, lastColumn);
		List<BomTemplate> bomTemplateList = new ArrayList<>();

		// Populate columns to java
		int rowNum = 1;
		for (List<String> row : rowList) {

			if (null == row.get(3) || row.get(3).isEmpty()) {

				System.out.println(
						"........BLANK MPN FOUND at row " + (rowNum + 1) + " in the File: " + bomTemplateFileName + "");
				System.out.println(
						"........IGNORING REMAINING rows. Please re-check the excel IF THIS IS NOT THE LAST ROW");
				System.out.println("..................................................................");
				System.out.println("..................................................................");
				break;
			}

			BomTemplate bomTemplate = new BomTemplate();
			bomTemplate.setFlexPartNo(row.get(FLEX_PART_NO_COLUMN_NUMBER));
			bomTemplate.setDescription(row.get(DESCRIPTION_COLUMN_NUMBER));
			bomTemplate.setManufacturer(row.get(MANUFACTURER_COLUMN_NUMBER));
			bomTemplate.setMcode(row.get(MCODE_COLUMN_NUMBER));
			bomTemplate.setMpn(row.get(MPN_COLUMN_NUMBER));
			bomTemplate.setQuantity(row.get(QUANTITY_COLUMN_NUMBER));
			bomTemplate.setEmailID(row.get(EMAILID_COLUMN_NUMBER));
			bomTemplate.setTelno(row.get(TEL_NO_COLUMN_NUMBER));
			bomTemplate.setAssemblyNo(row.get(ASSEMBLY_NO_COLUMN_NUMBER));
			bomTemplate.setCommodity(row.get(COMMODITY_COLUMN_NUMBER));

			bomTemplateList.add(bomTemplate);
			rowNum++;

		}

		return bomTemplateList;
	}
}
