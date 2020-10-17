package com.create.format.run;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ColorExample {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		try (OutputStream fileOut = new FileOutputStream("E:/Zoluxiones/Proyecto/3_Workspaces/create-excel-format-app/src/main/java/com/create/format/file/Javatpoint.xlsx")) {
			XSSFWorkbook wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet");
			
			
			Row row = sheet.createRow(1);
			CellStyle style = wb.createCellStyle();
			Cell cell = null;
			
			// Setting Foreground Color
			style = wb.createCellStyle();
			style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			cell = row.createCell(2);
			cell.setCellValue("A Technical Portal");
			cell.setCellStyle(style);
			
			cell = row.createCell(3);
			cell.setCellStyle(style);
			
			cell = row.createCell(4);
			cell.setCellStyle(style);
			
			
			
			wb.write(fileOut);
			wb.close();
			
			System.out.println("ok");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
