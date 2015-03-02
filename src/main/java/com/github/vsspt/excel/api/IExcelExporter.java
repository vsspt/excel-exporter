package com.github.vsspt.excel.api;

import java.util.List;
import java.util.Map;

import com.github.vsspt.excel.exception.ExcelReporterException;
import com.github.vsspt.excel.schema.ExcelReport;

public interface IExcelExporter {

	<T> byte[] listToExcel(List<T> data) throws ExcelReporterException;

	<T> byte[] listToExcel(List<T> data, List<String> ignore) throws ExcelReporterException;

	<T> byte[] listToExcel(List<T> data, List<String> ignore, Map<String, String> properties) throws ExcelReporterException;

	byte[] reportToExcel(ExcelReport report) throws ExcelReporterException;

}
