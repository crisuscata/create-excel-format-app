package com.create.format.run;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Template {

	public static void main(String[] args) {
		try {
			// Read the spreadsheet that needs to be updated
			FileInputStream fsIP = new FileInputStream(new File("E:\\Zoluxiones\\Proyecto\\3_Workspaces\\create-excel-format-app\\src\\main\\java\\com\\create\\format\\file\\template.xlsx"));
			// Access the workbook
			Workbook wb = new XSSFWorkbook(fsIP);
			// Access the worksheet, so that we can update / modify it.
			Sheet worksheet = wb.getSheetAt(0);
			// declare a Cell object
			Cell cell = null;
			// Access the second cell in second row to update the value
			cell = worksheet.getRow(0).getCell(0);
			// Get current cell value value and overwrite the value
			cell.setCellValue("OverRide existing value");
			// Close the InputStream
			fsIP.close();
			// Open FileOutputStream to write updates
			FileOutputStream output_file = new FileOutputStream(new File("E:\\\\Zoluxiones\\\\Proyecto\\\\3_Workspaces\\\\create-excel-format-app\\\\src\\\\main\\\\java\\\\com\\\\create\\\\format\\\\file\\\\templateOut.xlsx"));
			// write changes
			wb.write(output_file);
			// close the stream
			output_file.close();
			wb.close();
			
			System.out.println("OOOK");
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

}
