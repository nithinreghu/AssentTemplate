package com.flex.assenttemplate.util;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.flex.assenttemplate.dto.BomTemplate;
import com.flex.assenttemplate.dto.MstrDetails;

public class ValidationUtil {

	public static void validateBomTemplate(String bomTemplateFileName, Integer bomTemplateFirstRow, String mstrFileName,
			Integer mstrFirstRow) throws IOException {

		List<BomTemplate> bomTemplateData = getBomTemplateExcelData(bomTemplateFileName, bomTemplateFirstRow);
		Map<String, MstrDetails> mstrData = getMstrExcelData(mstrFileName, mstrFirstRow);

		validateMcode(bomTemplateData, mstrData);

	}

	private static Map<String, MstrDetails> getMstrExcelData(String mstrFileName, Integer mstrFirstRow)
			throws IOException {

		// Read excel
		List<List<String>> rowList = FileUtil.readExcel(mstrFileName, mstrFirstRow);
		Map<String, MstrDetails> mstrDetailsMap = new HashMap<>();

		// Populate columns to java
		for (List<String> row : rowList) {

			MstrDetails mstrDetails = new MstrDetails();
			mstrDetails.setGlobalMfgCodes(row.get(0));
			mstrDetails.setGlobalManufacturerName(row.get(1));
			mstrDetails.setObsolete(row.get(2));

			mstrDetailsMap.put(mstrDetails.getGlobalMfgCodes(), mstrDetails);

		}
		return mstrDetailsMap;
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

	private static void validateMcode(List<BomTemplate> bomTemplateData, Map<String, MstrDetails> mstrData) {
		// highlight mailID
		// Check if mcode in excel1 is available in excel2
		// ------if no, highlight mcode cell in excel1
		// --------------------------------------------------------------------------------------------
		// ------if yes,
		// -------- copy (mfrname, obsolute) value from excel2 to excel1
		// --------------------------------------------------------------------------------------------
		// -------- If mfr and manufacturer in excel1 are not equal, highlight
		// manufacturer cell in excel1

	}

}
