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

	private static void validateMcode(List<BomTemplate> bomTemplateList, Map<String, MstrDetails> mstrData) {

		// Get each row
		for (BomTemplate bomTemplate : bomTemplateList) {

			if (bomTemplate.getEmailID().toLowerCase().contains("@flex.com")
					|| bomTemplate.getEmailID().toLowerCase().contains("@bomcheck.com")) {
				
				// Highlight email column in bom template excel

			}

			// Get the values corresponding to mcode
			MstrDetails mstrDetails = mstrData.get(bomTemplate.getMcode());
			if (null != mstrDetails) {

				// String mfrname = mstrDetails.getGlobalManufacturerName();
				// String obsolute = mstrDetails.getObsolete();

				// Update (mfrname, obsolute) value in bom template excel

				if (!bomTemplate.getManufacturer().equals(mstrDetails.getGlobalManufacturerName())) {

					// highlight manufacturer cell in bom template excel
				}

			} else {

				// Highlight mcode cell in bom template excel
			}

		}

	}

}
