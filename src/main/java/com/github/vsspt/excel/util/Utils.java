package com.github.vsspt.excel.util;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.text.WordUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public final class Utils {

	private Utils() {}

	private static final short HEADER_SIZE = 14;

	private static final String FONT_NAME = "Arial";

	private static final String GET_FORMAT = "get%s";

	private static final String INSTANT_FORMAT = "yyMMddhhmmss";

	private static final String DATE_FORMAT = "yyyy/MM/dd";

	public static byte[] getBinaryExcelData(final Workbook workbook) throws IOException {
		if (workbook == null) {
			return null;
		}

		final ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
			workbook.write(bos);
			return bos.toByteArray();
		}
		finally {
			IOUtils.closeQuietly(bos);
		}
	}

	public static void saveWorksheet(final byte[] bytes, final String fullFileName) throws IOException {

		if (bytes != null) {
			final File file = new File(fullFileName);
			FileUtils.writeByteArrayToFile(file, bytes);
		}
	}

	public static void closeWorksheet(final Workbook workbook, final String workbookName) throws IOException {

		if (workbook != null) {
			FileOutputStream fileOut = null;
			try {
				fileOut = new FileOutputStream(workbookName);
				workbook.write(fileOut);
			}
			finally {
				IOUtils.closeQuietly(fileOut);
			}
		}
	}

	public static CellStyle createColumnHeaderCellStyle(final Workbook workbook) {
		final CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFillBackgroundColor(new HSSFColor.BLUE().getIndex());
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellStyle.setWrapText(Boolean.FALSE);
		cellStyle.setFont(getHeaderFont(workbook));

		return cellStyle;
	}

	public static Font getHeaderFont(final Workbook workbook) {
		final Font font = workbook.createFont();
		font.setFontHeightInPoints(HEADER_SIZE);
		font.setFontName(FONT_NAME);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		font.setColor(HSSFColor.WHITE.index);
		return font;
	}

	private static List<Field> getInheritedPrivateFields(final Class<?> type) {
		final List<Field> result = new ArrayList<Field>();

		Class<?> i = type;
		while (i != null && i != Object.class) {
			for (final Field field : i.getDeclaredFields()) {
				if (!field.isSynthetic()) {
					result.add(field);
				}
			}
			i = i.getSuperclass();
		}

		return result;
	}

	public static List<String> getFieldsForClass(final Class<?> clazz) throws Exception {

		final List<String> fieldNames = new LinkedList<String>();
		final List<Field> fields = getInheritedPrivateFields(clazz);

		for (final Field field : fields) {
			fieldNames.add(field.getName());
		}

		return fieldNames;
	}

	public static void setCellValue(final Cell cell, final Object value) {
		if (value != null) {
			if (value instanceof String) {
				cell.setCellValue((String) value);
			}
			else if (value instanceof Long) {
				cell.setCellValue((Long) value);
			}
			else if (value instanceof Integer) {
				cell.setCellValue((Integer) value);
			}
			else if (value instanceof Double) {
				cell.setCellValue((Double) value);
			}
			else if (value instanceof Date) {
				cell.setCellValue(formatDate((Date) value));
			}
			else if (value instanceof Float) {
				cell.setCellValue((Float) value);
			}
			else if (value instanceof Boolean) {
				cell.setCellValue((Boolean) value);
			}
		}
	}

	public static String capitalize(final String value) {
		return WordUtils.capitalize(value);
	}

	public static String getMethodName(final String value) {
		return String.format(GET_FORMAT, capitalize(value));
	}

	public static String getInstant() {
		return DateFormatUtils.format(new Date(), INSTANT_FORMAT);
	}

	public static String formatDate(final Date date) {
		return DateFormatUtils.format(date, DATE_FORMAT);
	}

}
