package com.ExcelParser.Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ExcelUtilities {
	
	private XSSFWorkbook workbook;
	private XSSFSheet activeSheet;
	
	public XSSFSheet loadWorkBook(String path) throws IOException {
		
		try {
			FileInputStream fileInputStream = new FileInputStream(new File(path));
			this.workbook = new XSSFWorkbook(fileInputStream);
			this.activeSheet = workbook.getSheetAt(0);
			return activeSheet;
			
		} catch (FileNotFoundException e) {
			
			System.out.println("File not found");
			return null;
		}

	}
	
	public String getCellValueInString(int columnId, int rowId) {
		Row row = activeSheet.getRow(rowId);
		Cell cell = row.getCell(columnId);
		String value = cell.getStringCellValue();
		return value;
	}
	
	public double getCellValueInDouble(int columnId, int rowId) {
		Row row = activeSheet.getRow(rowId);
		Cell cell = row.getCell(columnId);
		double value = cell.getNumericCellValue();
		return value;
	}
	
	public String getCellValueInDateFormat(int columnId, int rowId) {
		Row row = activeSheet.getRow(rowId);
		Cell cell = row.getCell(columnId);
		String value = cell.getDateCellValue().toString();
		return value;
	}
	
	
	public Row getEntireRow(int rowId) {
		Row row = activeSheet.getRow(rowId);
		return row;
	}
	
	public Column getEntireColumn(String columnId) {
		return null;
	}
}
