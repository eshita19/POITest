package com;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleGrid {

	public static void main(String[] args) {
		try (XSSFWorkbook workbook = new XSSFWorkbook();) {
			XSSFSheet sheet = workbook.createSheet();
			for (int i = 0; i < 4; i++) {
				XSSFRow row = sheet.createRow(i);
				for (int j = 0; j < 4; j++) {
					XSSFCell cell = row.createCell(j);
					cell.setCellValue(j);
				}
			}

			File file = new File("/Users/emathur/a.xlsx");
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
