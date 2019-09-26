package com;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellIconAndValTest {
	public static void main(String[] args) {
		try(Workbook workbook = new XSSFWorkbook();) {
			Sheet sheet = workbook.createSheet("Sheet1");

			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("       10000000");

			CreationHelper helper = workbook.getCreationHelper();
			Drawing drawing = sheet.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();

			anchor.setCol1(0);
			anchor.setRow1(0);
			anchor.setCol2(0);
			anchor.setRow2(0);

			// get the cell width of A1
			float cellWidthPx = sheet.getColumnWidthInPixels(0);
			System.out.println(cellWidthPx);

			// set wanted shape size
			int shapeWidthPx = 10;
			int shapeHeightPx = 10;

			// set the position of left edge as Dx1 in unit EMU
			anchor.setDx1(Math.round(1 * Units.EMU_PER_PIXEL));

			// set the position of right edge as Dx2 in unit EMU
			anchor.setDx2(Math.round((1 + shapeWidthPx) * Units.EMU_PER_PIXEL));

			// set upper padding
			int upperPaddingPx = 4;

			// set upper padding as Dy1 in unit EMU
			anchor.setDy1(upperPaddingPx * Units.EMU_PER_PIXEL);

			// set upper padding + shape height as Dy2 in unit EMU
			anchor.setDy2((upperPaddingPx + shapeHeightPx) * Units.EMU_PER_PIXEL);

			XSSFSimpleShape shape = ((XSSFDrawing) drawing).createSimpleShape((XSSFClientAnchor) anchor);
			shape.setShapeType(ShapeTypes.ELLIPSE);
			shape.setFillColor(255, 0, 0);

			FileOutputStream fileOut = new FileOutputStream("/Users/emathur/Downloads/poi_shapes_values.xlsx");
			workbook.write(fileOut);
			fileOut.close();

		} catch (IOException ioex) {
			ioex.printStackTrace();
		}
	}
}
