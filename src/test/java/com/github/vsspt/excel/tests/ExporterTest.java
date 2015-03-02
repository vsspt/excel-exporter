package com.github.vsspt.excel.tests;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.testng.AbstractTestNGSpringContextTests;
import org.testng.annotations.Test;

import com.github.vsspt.excel.api.IAnnotatedExcelExporter;
import com.github.vsspt.excel.api.IExcelExporter;
import com.github.vsspt.excel.exception.ExcelReporterException;
import com.github.vsspt.excel.schema.ExcelReport;
import com.github.vsspt.excel.util.Utils;

@ContextConfiguration(locations = { "classpath:exporterContext.xml" })
@Test
public class ExporterTest extends AbstractTestNGSpringContextTests {

	@Resource(name = "idAnnotatedExcelExporter")
	private IAnnotatedExcelExporter anService;

	@Resource(name = "idExcelExporter")
	private IExcelExporter service;

	@Test
	public void testAnnotatedExport() throws ExcelReporterException, IOException {

		final byte[] bytes = anService.listToExcel(getList());

		Utils.saveWorksheet(bytes, "testAnnotatedExport.xls");

		assert bytes != null;

	}

	@Test
	public void testExport() throws ExcelReporterException, IOException {

		final List<String> ignore = new ArrayList<String>();
		ignore.add("discard");
		final byte[] bytes = service.listToExcel(getList(), ignore);

		Utils.saveWorksheet(bytes, "testExport.xls");

		assert true;
	}

	@Test
	public void testGenericExport() throws ExcelReporterException, IOException {

		final byte[] bytes = service.reportToExcel(getExcelReport());

		Utils.saveWorksheet(bytes, "testGenericExport.xls");

		assert bytes != null;

	}

	private ExcelReport getExcelReport() {
		final ExcelReport report = new ExcelReport();
		final int nrows = 100;

		report.addHeader("Coluna 1");
		report.addHeader("Coluna 2");
		report.addHeader("Coluna 3");
		report.addHeader("Coluna 4");
		report.addHeader("Coluna 5");

		for (int i = 1; i <= nrows; i++) {

			final List<Object> row = new ArrayList<Object>();

			final Double col_1 = Double.valueOf(i * 2.5);
			final Long col_2 = Long.valueOf(i * 1);
			final String col_3 = "Coluna " + i;
			final Date col_4 = new Date();
			final Float col_5 = Float.valueOf((float) (i * 1.02));

			row.add(col_1);
			row.add(col_2);
			row.add(col_3);
			row.add(col_4);
			row.add(col_5);

			report.addRow(row);

		}

		return report;

	}

	private List<Model> getList() {
		final List<Model> list = new ArrayList<Model>();
		for (int i = 1; i <= 100; i++) {
			list.add(getModel(i));
		}
		return list;
	}

	private Model getModel(final int i) {
		final Model model = new Model();
		model.setAge(i);
		model.setName("Name " + i);
		model.setDiscard("XPTO");
		model.setMoney(i * 0.5d);

		return model;
	}
}
