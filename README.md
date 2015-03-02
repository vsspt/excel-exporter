# excel-exporter
===========
Exports Java objects to Microsoft Excel Reports.

Usage overview
===========
Offers 2 ways for exporting Java objects to Microsoft Excel Reports:
 - By annotating java objects using the annotation @ExcelColumn,
 - By creating and populating an ExcelReport object.
 
Annotation example
===========

Java class
```java
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

```

Exporting report
```java
	@Resource(name = "idAnnotatedExcelExporter")
	private IAnnotatedExcelExporter anService;
	
	public void testAnnotatedExport() throws ExcelReporterException, IOException {
		final byte[] bytes = anService.listToExcel(getObjectsList());
		Utils.saveWorksheet(bytes, "testAnnotatedExport.xls");	
	}
```



ExcelReport example
===========
```java

	public void testExcelReport() throws ExcelReporterException, IOException {

		final ExcelReport report = getExcelReport();
		final byte[] bytes = service.reportToExcel(report);

		Utils.saveWorksheet(bytes, "testGenericExport.xls");

		assert bytes != null;

	}	

	private ExcelReport getExcelReport() {
		final ExcelReport report = new ExcelReport();
		final int nrows = 100;

		report.addHeader("Column 1");
		report.addHeader("Column 2");
		report.addHeader("Column 3");
		report.addHeader("Column 4");
		report.addHeader("Column 5");

		for (int i = 1; i <= nrows; i++) {

			final List<Object> row = new ArrayList<Object>();

			final Double col_1 = Double.valueOf(i * 2.5);
			final Long col_2 = Long.valueOf(i * 1);
			final String col_3 = "Column " + i;
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
```

Maven Dependency - Service
===========
```xml
<dependency>
 <groupId>com.github.vsspt</groupId>
 <artifactId>excel-exporter</artifactId>
 <version>1.0.0</version>
</dependency>
```	


What's new
===========

Version 1.0.0
===========
- Initial deployment
