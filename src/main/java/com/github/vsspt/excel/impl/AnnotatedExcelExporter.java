package com.github.vsspt.excel.impl;

import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.vsspt.excel.annotation.ExcelColumn;
import com.github.vsspt.excel.annotation.ExcelReport;
import com.github.vsspt.excel.api.IAnnotatedExcelExporter;
import com.github.vsspt.excel.exception.ExcelReporterException;
import com.github.vsspt.excel.util.ExcelInfo;
import com.github.vsspt.excel.util.Utils;

public class AnnotatedExcelExporter implements IAnnotatedExcelExporter {

	private static final Logger LOG = LoggerFactory.getLogger(AnnotatedExcelExporter.class);

	private static final int ZERO = 0;

	private static final int ONE = 1;

	public AnnotatedExcelExporter() {
	}

	@Override
	public <T> byte[] listToExcel(final List<T> data) throws ExcelReporterException {

		if (data == null || data.isEmpty()) {
			LOG.debug("Nothing to export.");
			return null;
		}

		try {
			final ExcelInfo info = processAnnotations(data.get(ZERO));

			final HSSFWorkbook workbook = new HSSFWorkbook();

			final Sheet sheet = workbook.createSheet(data.get(ZERO).getClass().getName());

			final CellStyle columnHeaderCellStyle = Utils.createColumnHeaderCellStyle(workbook);
			int rowCount = ZERO;
			int columnCount = ZERO;

			Row row = sheet.createRow(rowCount++);
			for (final String labelName : info.getOrderLabels()) {
				final Cell cel = row.createCell(columnCount++);
				cel.setCellValue(labelName);
				cel.setCellStyle(columnHeaderCellStyle);
			}
			final Class<? extends Object> classz = data.get(ZERO).getClass();
			for (final T object : data) {
				row = sheet.createRow(rowCount++);

				columnCount = ZERO;

				for (final String label : info.getOrderLabels()) {
					final String methodName = info.getFieldLabelMap().get(label);
					final Cell cell = row.createCell(columnCount);
					final Method method = classz.getMethod(methodName);
					final Object value = method.invoke(object, (Object[]) null);

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

	private <T> ExcelInfo processAnnotations(final T object) {

		final ExcelInfo info = new ExcelInfo();

		final Class<?> clazz = object.getClass();
		final ExcelReport reportAnnotation = clazz.getAnnotation(ExcelReport.class);
		final String reportName = reportAnnotation.reportName();
		if (reportName == null || reportName.trim().length() < ONE) {
			LOG.error("Invalid Worksheet Name. Using [{}].", info.getWorkbookName());
		}
		else {
			info.setWorkbookName(reportAnnotation.reportName());
		}

		for (final Method method : clazz.getMethods()) {
			final ExcelColumn excelColumn = method.getAnnotation(ExcelColumn.class);
			if (excelColumn != null && !excelColumn.ignore()) {
				info.getFieldLabelMap().put(excelColumn.label(), method.getName());
				info.getOrderLabels().add(excelColumn.label());
			}
		}

		return info;
	}

}
