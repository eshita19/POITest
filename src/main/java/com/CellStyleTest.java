package com;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import javax.swing.GroupLayout.Alignment;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellStyleTest {
	public static void main(String[] args) {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("poi"));
			Map<String, Object> properties = new HashMap<String, Object>();
			
			// Adjust the column width to adjust content
			sheet.autoSizeColumn(1);
			CellStyle cellStyle = workbook.createCellStyle();
			
			cellStyle.setAlignment(HorizontalAlignment.CENTER);
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			//properties.put(CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
			//properties.put(CellUtil.ALIGNMENT, VerticalAlignment.CENTER);

			// wrap text, adjust column width to adjust content
			cellStyle.setWrapText(true);

			// Borders
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
			cellStyle.setBorderRight(BorderStyle.THIN);
			cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
			cellStyle.setBorderTop(BorderStyle.MEDIUM_DASHED);
			cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
			//properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);

			// Fill background color
			cellStyle.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());
			cellStyle.setFillPattern(FillPatternType.FINE_DOTS);

			// Create a new font and alter it.
			Font font = workbook.createFont();
			font.setColor(HSSFColor.RED.index);
			font.setFontHeightInPoints((short) 24);
			font.setFontName("Courier New");
			font.setItalic(true);
			font.setStrikeout(true);
			font.setBold(true);
			cellStyle.setFont(font);

			// Create a cell with a value and set style to it.
			for (int i = 0; i < 4; i++) {
				XSSFRow row = sheet.createRow(i);
				// Set Row height
				row.setHeightInPoints(12);
				XSSFCell cell = row.createCell(0);
				cell.setCellValue("Testing styles");
				cell.setCellStyle(cellStyle);
			}
			workbook.write(new FileOutputStream("/Users/emathur/Downloads/poi_cellstyle.xlsx"));
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
}
