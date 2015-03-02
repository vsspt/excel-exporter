package com.github.vsspt.excel.impl;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.vsspt.excel.api.IExcelExporter;
import com.github.vsspt.excel.exception.ExcelReporterException;
import com.github.vsspt.excel.schema.ExcelReport;
import com.github.vsspt.excel.util.Utils;

public class ExcelExporter implements IExcelExporter {

	private static final Logger LOG = LoggerFactory.getLogger(ExcelExporter.class);

	private static final int ZERO = 0;

	@Override
	public <T> byte[] listToExcel(final List<T> data) throws ExcelReporterException {
		return listToExcel(data, null, null);
	}

	@Override
	public <T> byte[] listToExcel(final List<T> data, final List<String> ignore) throws ExcelReporterException {
		return listToExcel(data, ignore, null);
	}

	@Override
	public <T> byte[] listToExcel(final List<T> data, final List<String> ignore, final Map<String, String> properties) throws ExcelReporterException {

		if (data == null || data.isEmpty() || data.get(ZERO) == null) {
			LOG.debug("Nothing to export.");
			return null;
		}

		try {

			final HSSFWorkbook workbook = new HSSFWorkbook();

			final Sheet sheet = workbook.createSheet(data.get(ZERO).getClass().getSimpleName());

			final List<String> fieldNames = Utils.getFieldsForClass(data.get(ZERO).getClass());

			// Create a row and put some cells in it. Rows are 0 based.
			int rowCount = ZERO;
			int columnCount = ZERO;

			final CellStyle columnHeaderCellStyle = Utils.createColumnHeaderCellStyle(workbook);

			Row row = sheet.createRow(rowCount++);
			for (final String fieldName : fieldNames) {
				if (!ignoreColumn(fieldName, ignore)) {
					final String columnName = getColumnName(fieldName, properties);
					final Cell cel = row.createCell(columnCount++);
					cel.setCellValue(columnName);
					cel.setCellStyle(columnHeaderCellStyle);
				}
			}
			final Class<? extends Object> classz = data.get(ZERO).getClass();
			for (final T object : data) {
				row = sheet.createRow(rowCount++);
				columnCount = ZERO;
				for (final String fieldName : fieldNames) {
					if (!ignoreColumn(fieldName, ignore)) {
						final Cell cell = row.createCell(columnCount);
						final Method method = classz.getMethod(Utils.getMethodName(fieldName));
						final Object value = method.invoke(object, (Object[]) null);
						Utils.setCellValue(cell, value);
						columnCount++;
					}
				}
			}

			adjustColumnsSize(sheet, columnCount);

			return Utils.getBinaryExcelData(workbook);

		}
		catch (final Exception ex) {
			LOG.error("Error on writeReportToExcel", ex);
			throw new ExcelReporterException(ex);
		}
	}

	@Override
	public byte[] reportToExcel(final ExcelReport report) throws ExcelReporterException {

		if (report == null || report.isEmpty()) {
			LOG.debug("Nothing to export.");
			return null;
		}

		LOG.debug(report.toString());

		try {
			final HSSFWorkbook workbook = new HSSFWorkbook();
			final Sheet sheet = workbook.createSheet(report.getSheetName());

			// Create a row and put some cells in it. Rows are 0 based.
			int rowCount = ZERO;
			int columnCount = ZERO;

			final CellStyle columnHeaderCellStyle = Utils.createColumnHeaderCellStyle(workbook);

			Row row = sheet.createRow(rowCount++);

			for (final String fieldName : report.getHeader()) {
				final Cell cel = row.createCell(columnCount++);
				cel.setCellValue(Utils.capitalize(fieldName));
				cel.setCellStyle(columnHeaderCellStyle);
			}

			for (final List<Object> columns : report.getRows()) {
				row = sheet.createRow(rowCount++);
				columnCount = ZERO;

				for (final Object value : columns) {
					final Cell cell = row.createCell(columnCount);

					Utils.setCellValue(cell, value);

					columnCount++;
				}
			}

			return Utils.getBinaryExcelData(workbook);

		}
		catch (final Exception ex) {
			LOG.error("Error on writeReportToExcel", ex);
			throw new ExcelReporterException(ex);
		}
	}

	private Boolean ignoreColumn(final String column, final List<String> ignore) {
		if (ignore == null) {
			return Boolean.FALSE;
		}

		return ignore.contains(column);

	}

	private String getColumnName(final String key, final Map<String, String> properties) {
		String name = key;
		if (properties != null && properties.containsKey(key)) {
			name = properties.get(key);
		}
		return Utils.capitalize(name);
	}

	private void adjustColumnsSize(final Sheet sheet, final int count) {
		for (int i = ZERO; i < count; i++) {
			sheet.autoSizeColumn(i);
		}
	}

}
