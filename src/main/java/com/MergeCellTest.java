package com;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergeCellTest {
	public static void main(String[] args) {
		try (Workbook wb = new XSSFWorkbook();) {
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(1);
			Cell cell = row.createCell(1);
			cell.setCellValue("Two cells have merged");
			// Merging cells by providing cell index
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));
			OutputStream fileOut = new FileOutputStream("/Users/emathur/Downloads/poi_merge_cells.xlsx");
			wb.write(fileOut);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
