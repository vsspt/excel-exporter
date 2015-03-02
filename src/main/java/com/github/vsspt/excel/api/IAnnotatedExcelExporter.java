package com.github.vsspt.excel.api;

import java.util.List;

import com.github.vsspt.excel.exception.ExcelReporterException;

public interface IAnnotatedExcelExporter {

	<T> byte[] listToExcel(List<T> data) throws ExcelReporterException;
}
