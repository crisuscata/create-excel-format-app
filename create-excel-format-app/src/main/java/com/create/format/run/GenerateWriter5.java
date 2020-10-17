package com.create.format.run;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Base64;
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

public class GenerateWriter5 {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		// Create a Workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		/*
		 * CreationHelper helps us create instances of various things like DataFormat,
		 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
		 */
		CreationHelper createHelper = workbook.getCreationHelper();

		/*** HOJA */
		Sheet sheet = workbook.createSheet("Fondos de Inversión");
		Sheet sheet2 = workbook.createSheet("Fondos Mutuos");

		// TEMP IMG
		addImage(workbook, sheet);
		addImage(workbook, sheet2);

		// COLUMN 0 ancho
		sizeFistColumn(sheet);
		sizeFistColumn(sheet2);

		// LINEA DIVISORA
		addLineDivide(workbook, sheet);
		addLineDivide(workbook, sheet2);
		
		// HIDE GRID LINE
		hideGridlines(sheet);
		hideGridlines(sheet2);

		addPrincipalTitle(workbook, sheet, "Fondos de Inversión");
		addPrincipalTitle(workbook, sheet2, "Fondos Mutuos");
		
		addTitlesHeaderGrid(workbook, sheet);
		addTitlesHeaderGrid(workbook, sheet2);

		
		List<Movimientos> movimientos = new ArrayList<>();
		movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A001", "Rescate",
				"Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A002", "Rescate",
				"Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A003", "Rescate",
				"Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A004", "Rescate",
				"Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A005", "Rescate",
				"Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		movimientos.add(new Movimientos("2019-03-27", "EMPRESA Préstamos Latinoamericanos A006", "Rescate",
				"Depósito de banco BANCO DE CRÉDITO DEL PERÚ", 0.029746, 107.4478461, -3.2, "DOL", "MUT", "-"));
		
		List<Movimientos> movimientos2 = new ArrayList<>();
		movimientos2.add(new Movimientos("2020-03-27", "EMPRESA Préstamos Latinoamericanos B001", "Rescate",
				"", 0.029746, 107.4478461, -3.2, "LOR", "RET", "-"));
		movimientos2.add(new Movimientos("2020-03-27", "EMPRESA Préstamos Latinoamericanos B002", "Rescate",
				"", 0.029746, 107.4478461, -3.2, "LOR", "RET", "-"));
		movimientos2.add(new Movimientos("2020-03-27", "EMPRESA Préstamos Latinoamericanos B003", "Rescate",
				"", 0.029746, 107.4478461, -3.2, "LOR", "RET", "-"));
		movimientos2.add(new Movimientos("2020-03-27", "EMPRESA Préstamos Latinoamericanos B004", "Rescate",
				"", 0.029746, 107.4478461, -3.2, "LOR", "RET", "-"));
		movimientos2.add(new Movimientos("2020-03-27", "EMPRESA Préstamos Latinoamericanos B005", "Rescate",
				"", 0.029746, 107.4478461, -3.2, "LOR", "RET", "-"));
		movimientos2.add(new Movimientos("2020-03-27", "EMPRESA Préstamos Latinoamericanos B006", "Rescate",
				"", 0.029746, 107.4478461, -3.2, "LOR", "RET", "-"));
		

		addBodyToGrid(workbook, createHelper, sheet, movimientos);
		addBodyToGrid(workbook, createHelper, sheet2, movimientos2);
		
		// Write the output to a file
		//FileOutputStream fileOut = new FileOutputStream(
		//		"E:/Zoluxiones/Proyecto/3_Workspaces/create-excel-format-app/src/main/java/com/create/format/file/poi-generated-file-EMPRESA.xlsx");

		 ByteArrayOutputStream bos = new ByteArrayOutputStream();
		 workbook.write(bos);
		 
		 byte[] bytesXLSX = bos.toByteArray();
		 
		// bos.close();
		 workbook.close();
		 
		 
		 String b64xlsx = Base64.getEncoder().encodeToString(bytesXLSX);
		 
		 System.out.println(b64xlsx);
		 //System.out.println("Excel Created");
		 
		 //InputStream is = this.getClass().getResourceAsStream("/icons/default_image.png");

	}

	private static void addImage(Workbook wb, Sheet sheet) {
		try {
			//InputStream inputStream = new FileInputStream(
			//		"/EMPRESA02.png");
			//Class cls = Class.forName("EMPRESAWriter5");
			//ClassLoader cLoader = cls.getClassLoader();
			//InputStream inputStream = cLoader.getResourceAsStream("/resources/img/logo_EMPRESA.png");
			InputStream inputStream = GenerateWriter5.class.getClassLoader().getResourceAsStream("img/logo_EMPRESA.png");
			
			byte[] bytes = IOUtils.toByteArray(inputStream);
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
			inputStream.close();
			CreationHelper helper = wb.getCreationHelper();
			Drawing<?> drawing = sheet.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();
			anchor.setCol1(7);
			anchor.setRow1(1);
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			double scale = 1;
			pict.resize(scale);
		} catch (Exception e) {
			e.getMessage();
		}

	}
	
	private static void sizeFistColumn(Sheet sheet) {
		sheet.setColumnWidth(0, 20 * 24);
	}
	
	private static void addLineDivide(XSSFWorkbook workbook, Sheet sheet) {
		Row row3 = sheet.createRow(3);
		row3.setHeight((short) 100);
		XSSFCellStyle cellStyleRow3 = workbook.createCellStyle();
		byte[] rgbRow3 = new byte[] { (byte) 28, (byte) 99, (byte) 158 };
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
	}
	
	private static void hideGridlines(Sheet sheet) {
		sheet.setDisplayGridlines(false);
	}
	
	private static void addPrincipalTitle(XSSFWorkbook workbook, Sheet sheet, String title) {
		Font headerFontTitle = workbook.createFont();
		headerFontTitle.setBold(true);
		headerFontTitle.setFontHeightInPoints((short) 22);
		headerFontTitle.setFontName("Arial");
		headerFontTitle.setBold(true);
		byte[] rgbTitle = new byte[] { (byte) 28, (byte) 99, (byte) 158 };
		XSSFFont xssfFont = (XSSFFont) headerFontTitle;
		xssfFont.setColor(new XSSFColor(rgbTitle, null));

		CellStyle headerCellStyleTitle = workbook.createCellStyle();
		headerCellStyleTitle.setFont(headerFontTitle);

		// Title word
		Row headerRowTitle = sheet.createRow(1);
		Cell cellTitle = headerRowTitle.createCell(1);
		cellTitle.setCellValue(title);
		sheet.addMergedRegion(CellRangeAddress.valueOf("B2:C2"));
		cellTitle.setCellStyle(headerCellStyleTitle);
	}
	
	private static void addTitlesHeaderGrid(XSSFWorkbook workbook, Sheet sheet) {
		 String[] columns = { "FECHA", "NOMBRE DEL FONDO", "DESCRIPCIÓN", "DETALLE", "N° CUOTAS",
				"VALOR CUOTA", "MONTO" };
		 
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
			int j = 0;
			for (int i = 0; i < columns.length; i++) {
				Cell cell = headerRow.createCell(++j);
				cell.setCellValue(columns[i]);
				cell.setCellStyle(headerCellStyle);
			}
			// Size
			headerRow.setHeight((short) 500);
	}
	
	private static void addBodyToGrid(XSSFWorkbook workbook, CreationHelper createHelper, Sheet sheet, List<Movimientos> movimientos) {
		Font dateFont = workbook.createFont();
		dateFont.setFontHeightInPoints((short) 10);
		dateFont.setFontName("Arial");
		byte[] rgbDate = new byte[] { (byte) 128, (byte) 128, (byte) 128 };
		XSSFFont xssfFontDate = (XSSFFont) dateFont;
		xssfFontDate.setColor(new XSSFColor(rgbDate, null));

		byte[] rgbBorder = new byte[] { (byte) 216, (byte) 216, (byte) 216 };

		XSSFCellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));
		dateCellStyle.setFont(dateFont);
		dateCellStyle.setBorderBottom(BorderStyle.MEDIUM);
		dateCellStyle.setBottomBorderColor(new XSSFColor(rgbBorder, null));
		dateCellStyle.setAlignment(HorizontalAlignment.LEFT);
		dateCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		// Create Cell Style for formatting Monto
		XSSFCellStyle amountCellStyle = workbook.createCellStyle();
		amountCellStyle.setFont(dateFont);
		amountCellStyle.setBorderBottom(BorderStyle.MEDIUM);
		amountCellStyle.setBottomBorderColor(new XSSFColor(rgbBorder, null));
		amountCellStyle.setAlignment(HorizontalAlignment.LEFT);
		amountCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		Font otherRowFont = workbook.createFont();
		otherRowFont.setFontHeightInPoints((short) 10);
		otherRowFont.setFontName("Arial");
		byte[] rgbOther = new byte[] { (byte) 28, (byte) 99, (byte) 158 };
		XSSFFont xssfFontOther = (XSSFFont) otherRowFont;
		xssfFontOther.setColor(new XSSFColor(rgbOther, null));

		XSSFCellStyle otherCellStyle = workbook.createCellStyle();
		otherCellStyle.setFont(otherRowFont);
		otherCellStyle.setBorderBottom(BorderStyle.MEDIUM);
		otherCellStyle.setBottomBorderColor(new XSSFColor(rgbBorder, null));
		otherCellStyle.setAlignment(HorizontalAlignment.LEFT);
		otherCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		otherCellStyle.setWrapText(true); // AJUSTE AUTOMATICO

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
			amountCell.setCellValue(objMovim.getNumeroDeCuotas() * objMovim.getValorCuota());
			amountCell.setCellStyle(amountCellStyle);

			row.setHeight((short) 800);
		}

		// Resize all columns to fit the content size
		int totalColums = 7;
		for (int i = 0; i < totalColums; i++) {
			sheet.autoSizeColumn(i);
		}
		
		sheet.setColumnWidth(1, 70 * 45);
		sheet.setColumnWidth(2, 90 * 90);
		//sheet.setColumnWidth(4, 90 * 120);
		sheet.setColumnWidth(7, 90 * 40);
	}

}
