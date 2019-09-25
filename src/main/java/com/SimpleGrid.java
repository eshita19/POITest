package com;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleGrid {

	public static void main(String[] args) {
		try (XSSFWorkbook workbook = new XSSFWorkbook();) {
			XSSFSheet sheet = workbook.createSheet();
			// Create a new font and alter it.
			XSSFFont font = workbook.createFont();
			font.setFontHeightInPoints((short) 30);
			font.setFontName("IMPACT");
			font.setItalic(true);
			// font.setColor(XSSFColor.from(CTColor.EQUAL , map));

			// Set font into style
			XSSFCellStyle style = workbook.createCellStyle();
			style.setFont(font);

			// Create a cell with a value and set style to it.
			for (int i = 0; i < 4; i++) {
				XSSFRow row = sheet.createRow(i);
				for (int j = 0; j < 4; j++) {
					XSSFCell cell = row.createCell(j);
					cell.setCellValue(j);
					// cell.setCellStyle(style);
				}
			}

			File file = new File("/Users/emathur/Downloads/poi.xlsx");
			try {
				workbook.write(new FileOutputStream(file));
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}
}
