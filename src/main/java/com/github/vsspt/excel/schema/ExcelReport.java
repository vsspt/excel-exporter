package com.github.vsspt.excel.schema;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import com.github.vsspt.excel.util.Utils;

public class ExcelReport {

	private static final String FORMAT = "%s_%s.%s";

	private String sheetName = "Sheet 1";

	private String fileName = "ExcelReport";

	private String fileExtension = "xls";

	private final List<String> header = new ArrayList<String>();

	private final List<List<Object>> rows = new ArrayList<List<Object>>();

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(final String sheetName) {
		this.sheetName = sheetName;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(final String fileName) {
		this.fileName = fileName;
	}

	public String getFileExtension() {
		return fileExtension;
	}

	public void setFileExtension(final String fileExtension) {
		this.fileExtension = fileExtension;
	}

	public void addHeader(final String value) {
		if (!header.contains(value)) {
			header.add(value);
		}
	}

	public void addRow(final List<Object> list) {
		rows.add(Collections.unmodifiableList(list));
	}

	public List<String> getHeader() {
		return Collections.unmodifiableList(header);
	}

	public List<List<Object>> getRows() {
		return Collections.unmodifiableList(rows);
	}

	public String getReportFullName() {
		return String.format(FORMAT, fileName, Utils.getInstant(), fileExtension);
	}

	public boolean isEmpty() {
		return rows.isEmpty();
	}

	@Override
	public String toString() {
		return String.format("Report Name [%s], Headers [%s].", getReportFullName(), header.toString());
	}

}
