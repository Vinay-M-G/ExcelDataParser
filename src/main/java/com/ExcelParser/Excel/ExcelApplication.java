package com.ExcelParser.Excel;

import java.io.IOException;
import java.util.List;

public class ExcelApplication 
{
    public static void main( String[] args )
    {
       
    	String dataInputPath = "E:\\Technical Stuff\\PracticeSession\\Source Book.xlsx";
    	String dataOutputPath = "E:\\Technical Stuff\\PracticeSession\\Destination Book.xlsx";
    	
    	ExcelDataHandler excelDataHandler = new ExcelDataHandler();
    	
    	try {
			List<ExcelDataAttributes> data = excelDataHandler.getData(dataInputPath);
			excelDataHandler.addDataToNewExcel(dataOutputPath, data);
			
		} catch (IOException e) {
			
			e.printStackTrace();
		}
    }
}
