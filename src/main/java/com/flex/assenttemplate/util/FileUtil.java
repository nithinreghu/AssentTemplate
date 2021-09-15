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

public class FileUtil {

	public static List<List<String>> readExcel(String filename, int firstRow) throws IOException {

		List<List<String>> rowList = new ArrayList<>();
		Workbook workbook = null;
		try {
			List<String> columnList = null;
			workbook = WorkbookFactory.create(new File(filename));
			Sheet sheet = workbook.getSheetAt(0);

			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			for (int i = firstRow - 1; i <= sheet.getLastRowNum(); i++) {

				columnList = new ArrayList<>();
				Row row = sheet.getRow(i);

				for (int j = 0; j < row.getLastCellNum(); j++) {
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
}
