package com.football;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToInsertSql {

	public static final String OFFICE_EXCEL_XLS = "xls";
	public static final String OFFICE_EXCEL_XLSX = "xlsx";

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		if (args.length != 2) {
			System.out.println("参数错了： 第一个参数：excel文件位置，第二个参数：表名称");
		}
		String path = args[0];
		String tableName = args[1];

		Workbook workbook = getWorkbook(path);
		if (workbook == null) {
			return;
		}
		Sheet sheet = workbook.getSheetAt(0);
		String headStr = "INSERT INTO " + tableName + " (";
		int totalRow = sheet.getLastRowNum();
		for (int i = 0; i < totalRow; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			// title
			if (i == 0) {
				headStr = readTitle(headStr, row);
				continue;
			}
			String insertStr = headStr;
			boolean isAdd = true;
			// context
			for (int j = 0; j < row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					isAdd = false;
					continue;
				}
				switch (cell.getCellTypeEnum()) {
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						Date date = cell.getDateCellValue();
						DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
						insertStr = insertStr + "'" + formater.format(date) + "' ,";
					} else if (String.valueOf(cell.getNumericCellValue()).contains(".")) {
						DecimalFormat df = new DecimalFormat("#");
						insertStr = insertStr + df.format(cell.getNumericCellValue()) + " ,";
					} else {
						insertStr = insertStr + (cell + "").trim() + " ,";
					}
					continue;
				case STRING:
					insertStr = insertStr + "'" + cell.getStringCellValue() + "' ,";
					continue;
				default:
					continue;
				}
			}
			if (isAdd) {
				insertStr = insertStr.substring(0, insertStr.length() - 1) + " );";
				System.out.println(insertStr);
			}
		}

	}

	private static String readTitle(String headStr, Row row) {
		for (int j = 0; j < row.getLastCellNum(); j++) {
			headStr = headStr + row.getCell(j).getStringCellValue() + ",";
		}
		headStr = headStr.substring(0, headStr.length() - 1);
		headStr = headStr + " ) VALUES (";
		return headStr;
	}

	/**
	 * 根据文件路径获取Workbook对象
	 * 
	 * @param filepath 文件全路径
	 */
	public static Workbook getWorkbook(String filepath)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		InputStream is = null;
		Workbook wb = null;
		if (StringUtils.isBlank(filepath)) {
			throw new IllegalArgumentException("文件路径不能为空");
		} else {
			String suffiex = getSuffiex(filepath);
			if (StringUtils.isBlank(suffiex)) {
				throw new IllegalArgumentException("文件后缀不能为空");
			}
			if (OFFICE_EXCEL_XLS.equals(suffiex) || OFFICE_EXCEL_XLSX.equals(suffiex)) {
				try {
					is = new FileInputStream(filepath);
					wb = WorkbookFactory.create(is);
				} finally {
					if (is != null) {
						is.close();
					}
					if (wb != null) {
						wb.close();
					}
				}
			} else {
				throw new IllegalArgumentException("该文件非Excel文件");
			}
		}
		return wb;
	}

	/**
	 * 获取后缀
	 * 
	 * @param filepath filepath 文件全路径
	 */
	private static String getSuffiex(String filepath) {
		if (StringUtils.isBlank(filepath)) {
			return "";
		}
		int index = filepath.lastIndexOf(".");
		if (index == -1) {
			return "";
		}
		return filepath.substring(index + 1, filepath.length());
	}


}