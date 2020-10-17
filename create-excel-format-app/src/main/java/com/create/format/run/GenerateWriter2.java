package com.create.format.run;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.create.format.entity.Movimientos;

public class GenerateWriter2 {
	
	private static String[] columns = { "FECHA", "NOMBRE DEL FONDO", "DESCRIPCIÓN", "DETALLE", "N° CUOTAS", "VALOR CUOTA", "MONTO" };
	private static List<Movimientos> movimientos = new ArrayList<>();
	
	// Initializing employees data to insert into the excel file
		static {
			movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A001", "Rescate", "Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
			movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A002", "Rescate", "Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
			movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A003", "Rescate", "Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
			movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A004", "Rescate", "Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
			movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A005", "Rescate", "Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
			movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A006", "Rescate", "Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		}

		public static void main(String[] args) throws IOException, InvalidFormatException {
			// Create a Workbook
			XSSFWorkbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

			/*
			 * CreationHelper helps us create instances of various things like DataFormat,
			 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
			 */
			CreationHelper createHelper = workbook.getCreationHelper();

			/*** TITLE */
			Sheet sheet = workbook.createSheet("Fondos de Inversión");
			
			//TEMP IMG
			
			addImage(workbook, sheet);
			//
			
			//COLUMN 0 ancho
			sheet.setColumnWidth(0, 20 * 24);
			
			
			//sheet.setColumnWidth(3, 20 * 24);
			//
			
			//LINEA DIVISORA
			Row row3 = sheet.createRow(3);
			row3.setHeight((short)100);
			XSSFCellStyle cellStyleRow3 = workbook.createCellStyle();
			byte[] rgbRow3 = new byte[]{(byte)28, (byte)99, (byte)158};
			cellStyleRow3.setFillForegroundColor(new XSSFColor(rgbRow3, null));
			cellStyleRow3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			Cell cellRow = null;
			
			cellRow = row3.createCell(1);
			cellRow.setCellStyle(cellStyleRow3);
			
			cellRow = row3.createCell(2);
			cellRow.setCellStyle(cellStyleRow3);
			
			cellRow = row3.createCell(3);
			cellRow.setCellStyle(cellStyleRow3);
			
			cellRow = row3.createCell(4);
			cellRow.setCellStyle(cellStyleRow3);
			
			cellRow = row3.createCell(5);
			cellRow.setCellStyle(cellStyleRow3);
			
			cellRow = row3.createCell(6);
			cellRow.setCellStyle(cellStyleRow3);
			
			cellRow = row3.createCell(7);
			cellRow.setCellStyle(cellStyleRow3);
			//
			
			
			//HIDE GRID LINE
			sheet.setDisplayGridlines(false);
			
			Font headerFontTitle = workbook.createFont();
			headerFontTitle.setBold(true);
			headerFontTitle.setFontHeightInPoints((short) 22);
			headerFontTitle.setFontName("Arial");
			headerFontTitle.setBold(true);
			byte[] rgbTitle = new byte[]{(byte)28, (byte)99, (byte)158};
			XSSFFont xssfFont = (XSSFFont)headerFontTitle;
			xssfFont.setColor(new XSSFColor(rgbTitle, null));
			
			
			CellStyle headerCellStyleTitle = workbook.createCellStyle();
			headerCellStyleTitle.setFont(headerFontTitle);
			
			
			// Title word
			Row headerRowTitle = sheet.createRow(1);
			Cell cellTitle = headerRowTitle.createCell(1);
			cellTitle.setCellValue("Fondos de Inversión");
			sheet.addMergedRegion(CellRangeAddress.valueOf("B2:C2"));
			cellTitle.setCellStyle(headerCellStyleTitle);
			
			
			
			/**
			 * HEADER TITLE OF ROWS
			 * ***/
			// Create a Font for styling header cells
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 10);
			headerFont.setFontName("Arial");
			headerFont.setColor(IndexedColors.BLACK.getIndex());

			// Create a CellStyle with the font
			XSSFCellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);
			headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
			headerCellStyle.setAlignment(HorizontalAlignment.LEFT);
			headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

			// Create a Row
			Row headerRow = sheet.createRow(5);

			// Create cells
			int j=0;
			for (int i = 0; i < columns.length; i++) {
				Cell cell = headerRow.createCell(++j);
				cell.setCellValue(columns[i]);
				cell.setCellStyle(headerCellStyle);
			}
			//Size
			headerRow.setHeight((short)500);
			
			//
			
			
			/**BODY**/
			
			//ESTILOS
			// Create Cell Style for formatting Date
			Font dateFont = workbook.createFont();
			dateFont.setFontHeightInPoints((short) 10);
			dateFont.setFontName("Arial");
			byte[] rgbDate = new byte[]{(byte)128, (byte)128, (byte)128};
			XSSFFont xssfFontDate = (XSSFFont)dateFont;
			xssfFontDate.setColor(new XSSFColor(rgbDate, null));
			
			byte[] rgbBorder = new byte[]{(byte)216, (byte)216, (byte)216};
			
			XSSFCellStyle dateCellStyle = workbook.createCellStyle();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));
			dateCellStyle.setFont(dateFont);
			dateCellStyle.setBorderBottom(BorderStyle.MEDIUM);
			dateCellStyle.setBottomBorderColor(new XSSFColor(rgbBorder, null));
			dateCellStyle.setAlignment(HorizontalAlignment.LEFT);
			dateCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			
			//  Create Cell Style for formatting Monto
			XSSFCellStyle  amountCellStyle = workbook.createCellStyle();
			amountCellStyle.setFont(dateFont);
			amountCellStyle.setBorderBottom(BorderStyle.MEDIUM);
			amountCellStyle.setBottomBorderColor(new XSSFColor(rgbBorder, null));
			amountCellStyle.setAlignment(HorizontalAlignment.LEFT);
			amountCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			
			Font otherRowFont = workbook.createFont();
			otherRowFont.setFontHeightInPoints((short) 10);
			otherRowFont.setFontName("Arial");
			byte[] rgbOther = new byte[]{(byte)28, (byte)99, (byte)158};
			XSSFFont xssfFontOther = (XSSFFont)otherRowFont;
			xssfFontOther.setColor(new XSSFColor(rgbOther, null));
			
			XSSFCellStyle otherCellStyle = workbook.createCellStyle();
			otherCellStyle.setFont(otherRowFont);
			otherCellStyle.setBorderBottom(BorderStyle.MEDIUM);
			otherCellStyle.setBottomBorderColor(new XSSFColor(rgbBorder, null));
			otherCellStyle.setAlignment(HorizontalAlignment.LEFT);
			otherCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			otherCellStyle.setWrapText(true); //AJUSTE AUTOMATICO

			// Create Other rows and cells with BODY
			int rowNum = 6;
			for (Movimientos objMovim : movimientos) {
				Row row = sheet.createRow(rowNum++);

				Cell dateCell = row.createCell(1);
				dateCell.setCellValue(objMovim.getFecha());
				dateCell.setCellStyle(dateCellStyle);
				
				Cell cell = null;
				
				cell = row.createCell(2);
				cell.setCellValue(objMovim.getFondo());
				cell.setCellStyle(otherCellStyle);
				
				cell = row.createCell(3);
				cell.setCellValue(objMovim.getDescripcion());
				cell.setCellStyle(otherCellStyle);
				
				cell = row.createCell(4);
				cell.setCellValue(objMovim.getDetalle());
				cell.setCellStyle(otherCellStyle);
				
				cell = row.createCell(5);
				cell.setCellValue(objMovim.getNumeroDeCuotas());
				cell.setCellStyle(otherCellStyle);
				
				cell = row.createCell(6);
				cell.setCellValue(objMovim.getValorCuota());
				cell.setCellStyle(otherCellStyle);
				
				Cell amountCell = row.createCell(7);
				amountCell.setCellValue(objMovim.getNumeroDeCuotas()*objMovim.getValorCuota());
				amountCell.setCellStyle(amountCellStyle);
				
				
				row.setHeight((short)800);
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < columns.length; i++) {
				sheet.autoSizeColumn(i);
			}
			
			sheet.setColumnWidth(2,90*90);
			sheet.setColumnWidth(4,90*120);
			sheet.setColumnWidth(7,90*40);
			
			
			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream("E:/Zoluxiones/Proyecto/3_Workspaces/create-excel-format-app/src/main/java/com/create/format/file/poi-generated-file-EMPRESA.xlsx");
			
			//ByteArrayOutputStream bos = new ByteArrayOutputStream();
			
			workbook.write(fileOut); 
			
			fileOut.close();

			// Closing the workbook
			workbook.close();
			
			//byte[] bytes = bos.toByteArray(); 
			
			System.out.println("Excel Created");
			
		}
		
		
		
		
		
		
		private static void addImage(Workbook wb , Sheet sheet) {
			
			try {
				 
				   //FileInputStream obtains input bytes from the image file
				   InputStream inputStream = new FileInputStream("E:/Zoluxiones/Proyecto/3_Workspaces/create-excel-format-app/src/main/java/com/create/format/file/EMPRESA02.png");
				   //Get the contents of an InputStream as a byte[].
				   byte[] bytes = IOUtils.toByteArray(inputStream);
				   //Adds a picture to the workbook
				   int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
				   //close the input stream
				   inputStream.close();
				 
				   //Returns an object that handles instantiating concrete classes
				   CreationHelper helper = wb.getCreationHelper();
				 
				   //Creates the top-level drawing patriarch.
				   Drawing<?> drawing = sheet.createDrawingPatriarch();
				 
				   //Create an anchor that is attached to the worksheet
				   ClientAnchor anchor = helper.createClientAnchor();
				   //set top-left corner for the image
				   anchor.setCol1(7);
				   anchor.setRow1(1);
				 
				   //Creates a picture
				   Picture pict = drawing.createPicture(anchor, pictureIdx);
				   //Reset the image to the original size
				   double scale = 1;
				   pict.resize(scale);
				
				
			} catch (Exception e) {
			}
			
			
		}
		
		
		
		
		
		
		
		
		
		
		
		

}
