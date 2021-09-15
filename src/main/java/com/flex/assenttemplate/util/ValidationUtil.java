package com.flex.assenttemplate.util;

import java.io.IOException;
import java.util.List;

import com.flex.assenttemplate.dto.BomTemplate;
import com.flex.assenttemplate.dto.MstrDetails;

public class ValidationUtil {

	public static void validateBomTemplate(String bomTemplateFileName, Integer bomTemplateFirstRow, String mstrFileName,
			Integer mstrFirstRow) throws IOException {

		List<BomTemplate> bomTemplateData = getBomTemplateExcelData(bomTemplateFileName, bomTemplateFirstRow);
		List<MstrDetails> mstrData = getMstrExcelData(mstrFileName, mstrFirstRow);

		validateMcode(bomTemplateData, mstrData);

	}

	private static List<MstrDetails> getMstrExcelData(String mstrFileName, Integer mstrFirstRow) throws IOException {
		List<List<String>> rowList = FileUtil.readExcel(mstrFileName, mstrFirstRow);
		return null;
	}

	private static List<BomTemplate> getBomTemplateExcelData(String bomTemplateFileName, Integer bomTemplateFirstRow)
			throws IOException {
		List<List<String>> rowList = FileUtil.readExcel(bomTemplateFileName, bomTemplateFirstRow);
		return null;
	}

	private static void validateMcode(List<BomTemplate> bomTemplateData, List<MstrDetails> mstrData) {
		// highlight mailID
		// Check if mcode in excel1 is available in excel2
		// ------if no, highlight mcode cell in excel1
		// --------------------------------------------------------------------------------------------
		// ------if yes,
		// -------- copy (mfrname, obsolute) value from excel2 to excel1
		// --------------------------------------------------------------------------------------------
		// -------- If mfr and manufacturer in excel1 are not equal, highlight manufacturer cell in excel1

	}

}
