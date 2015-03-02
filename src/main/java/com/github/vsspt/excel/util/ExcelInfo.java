package com.github.vsspt.excel.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelInfo {

	private String workbookName = "workbook.xls";

	private final Map<String, String> fieldLabelMap = new HashMap<String, String>();

	private final List<String> orderLabels = new ArrayList<String>();

	public String getWorkbookName() {
		return workbookName;
	}

	public void setWorkbookName(final String workbookName) {
		this.workbookName = workbookName;
	}

	public Map<String, String> getFieldLabelMap() {
		return fieldLabelMap;
	}

	public List<String> getOrderLabels() {
		return orderLabels;
	}

}
