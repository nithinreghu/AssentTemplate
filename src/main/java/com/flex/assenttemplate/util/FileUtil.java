package com.flex.assenttemplate.util;

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
		for (List<String> row : rowList) {

			BomTemplate bomTemplate = new BomTemplate();
			bomTemplate.setFlexPartNo(row.get(0));
			bomTemplate.setDescription(row.get(1));
			bomTemplate.setManufacturer(row.get(2));
			bomTemplate.setMcode(row.get(3));
			bomTemplate.setMpn(row.get(4));
			bomTemplate.setQuantity(row.get(5));
			bomTemplate.setEmailID(row.get(6));
			bomTemplate.setTelno(row.get(7));
			bomTemplate.setAssemblyNo(row.get(8));

			bomTemplateList.add(bomTemplate);

		}

		return bomTemplateList;
	}
}
