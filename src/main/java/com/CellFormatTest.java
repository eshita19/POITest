package com;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellFormatTest {
	public static void main(String[] args) {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("poi"));
			DataFormat format = workbook.createDataFormat();
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setDataFormat(format.getFormat("#,##0.0000"));

			// Create a cell with a value and set style to it.
			for (int i = 0; i < 4; i++) {
				XSSFRow row = sheet.createRow(i);
				int j = 0;
				XSSFCell cell = row.createCell(j++);
				cell.setCellValue(99999);
				cell.setCellStyle(cellStyle);

			}
			workbook.write(new FileOutputStream("/Users/emathur/Downloads/poi_data_format.xlsx"));
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
}
