package com.github.vsspt.excel.tests;

import com.github.vsspt.excel.annotation.ExcelColumn;
import com.github.vsspt.excel.annotation.ExcelReport;

@ExcelReport(reportName = "modeReport.xls")
public class Model {

	private String name;

	private int age;

	private double money;

	private String discard;

	@ExcelColumn(label = "Name")
	public String getName() {
		return name;
	}

	public void setName(final String name) {
		this.name = name;
	}

	@ExcelColumn(label = "Age")
	public int getAge() {
		return age;
	}

	public void setAge(final int age) {
		this.age = age;
	}

	@ExcelColumn(label = "Money")
	public double getMoney() {
		return money;
	}

	public void setMoney(final double money) {
		this.money = money;
	}

	@ExcelColumn(ignore = true)
	public String getDiscard() {
		return discard;
	}

	public void setDiscard(final String discard) {
		this.discard = discard;
	}

}
