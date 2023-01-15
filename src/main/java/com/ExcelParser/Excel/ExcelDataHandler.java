package com.ExcelParser.Excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Busnisess Logic
 */
public class ExcelDataHandler {
	
	public List<ExcelDataAttributes> getData(String path) throws IOException{
		
		ExcelUtilities excelUtilities = new ExcelUtilities();
		XSSFSheet sheet = excelUtilities.loadWorkBook(path);
		List<ExcelDataAttributes> dataList = new ArrayList<>();
		
		if(sheet != null) {
			
			for(int rowIndex = 0; rowIndex < 15; rowIndex++) {
				
				Row row = sheet.getRow(rowIndex);
				String transcationDate = "";
				
				try {
					transcationDate = row.getCell(0).getDateCellValue().toString();
					
				}catch(Exception ex) {
					
				}
				
				
				if(!transcationDate.isEmpty()) {
					
					ExcelDataAttributes excelDataAttributes = new ExcelDataAttributes();
					
					excelDataAttributes.setPartyName(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getPartyName());
					excelDataAttributes.setInvoiceNo(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getInvoiceNo());
					excelDataAttributes.setInvoiceAmount(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getInvoiceAmount());
					excelDataAttributes.setPvn(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getPvn());
					excelDataAttributes.setDeductedAmount(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getDeductedAmount());
					excelDataAttributes.setDeductionReason(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getDeductionReason());
					excelDataAttributes.setPaymentDate(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getPaymentDate());
					excelDataAttributes.setBankReference(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getBankReference());
					excelDataAttributes.setPaymentType(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getPaymentType());
					excelDataAttributes.setTransferAmount(excelUtilities, rowIndex, 0);
					System.out.println(excelDataAttributes.getTransferAmount());
					
					dataList.add(excelDataAttributes);
					
					System.out.println("======================================");
					
				}
			}
			
			
		}
		
		return dataList;
		
	}
	
	
	public void addDataToNewExcel(String path, List<ExcelDataAttributes> dataList) {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("filtered Data");
		
		int rownum = 0;
		Row headerRow = sheet.createRow(rownum);
		
		ExcelDataAttributes excelDataAttributes = new ExcelDataAttributes();
		List<String> dataHeader = excelDataAttributes.getDataAttributes();
		
		for(int index = 0 ; index < dataHeader.size(); index++) {
			headerRow.createCell(index).setCellValue(dataHeader.get(index));
		}
		
		rownum++;
		
		for(ExcelDataAttributes element : dataList) {
			int columnId = 0;
			Row row = sheet.createRow(rownum);
			
			row.createCell(columnId).setCellValue(element.getPartyName());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getInvoiceNo());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getInvoiceAmount());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getPvn());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getDeductedAmount());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getDeductionReason());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getPaymentDate());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getBankReference());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getPaymentType());
			columnId++;
			
			row.createCell(columnId).setCellValue(element.getTransferAmount());
			columnId++;
			
			rownum++;
			
		}
		
		try {
			 FileOutputStream out = new FileOutputStream(new File(path));
			 workbook.write(out);
			 workbook.close();
			 out.close();
			 System.out.println("File Created Successfully");
			
		}catch(Exception ex) {
			
			System.out.println(ex.toString());
			
		}
		
		
		
	}
	
}
