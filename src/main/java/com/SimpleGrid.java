package com;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleGrid {
	public static void main(String[] args) {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			CreationHelper createHelper = workbook.getCreationHelper();
			XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("poi"));
			// Create a cell with a value and set style to it.
			for (int i = 0; i < 4; i++) {
				XSSFRow row = sheet.createRow(i);
				int j = 0;
				// Integer
				CellUtil.createCell(row, j++, "1");
				// Float
				CellUtil.createCell(row, j++, "1.2");
				// Rich text string
				row.createCell(j++).setCellValue(createHelper.createRichTextString("This is a string"));
				// Boolean
				CellUtil.createCell(row, j++, "true");
				row.createCell(j++).setCellValue(true);
				
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
				// Date
				XSSFCell dateCell = row.createCell(j++);
				dateCell.setCellStyle(cellStyle);
				dateCell.setCellValue(new Date());
				// Calendar
				XSSFCell calendarCell = row.createCell(j++);
				calendarCell.setCellStyle(cellStyle);
				calendarCell.setCellValue(Calendar.getInstance());
				// Error
				XSSFCell errorCell = row.createCell(j++);
				errorCell.setCellType(CellType.ERROR);
				errorCell.setCellValue("error");
			}
			workbook.write(new FileOutputStream("/Users/emathur/Downloads/poi_datatypes.xlsx"));
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
}
