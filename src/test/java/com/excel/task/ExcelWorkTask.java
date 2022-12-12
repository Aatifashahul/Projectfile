package com.excel.task;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelWorkTask {
     @Test
	public void exceltoread() throws Exception   {
    	 File f = new File(System.getProperty("user.dir") +
    			 "\\src\\test\\resources\\Demo Sheet.xlsx");
    			 FileInputStream input = new FileInputStream(f);
    			 XSSFWorkbook workbook = new XSSFWorkbook(input);
    			 XSSFSheet sheet = workbook.getSheet("Demo page");
    			 int totalRows = sheet.getPhysicalNumberOfRows();
    			 
    			 for (int i = 0; i < totalRows; i++) {
    			 XSSFRow row = sheet.getRow(i);
    			 int totalcells = row.getPhysicalNumberOfCells();
    			 for (int j = 0; j < totalcells; j++) {
    			 XSSFCell cell = row.getCell(j);
    			 
    			 if (cell.getCellType() == CellType.NUMERIC) {
    			 double numericCellValue =
    			 cell.getNumericCellValue();
    			 System.out.println(numericCellValue + " ");
    			 } else {
    			 String stringCellValue = cell.getStringCellValue();
    			 System.out.println(stringCellValue + " ");
    			 } }
    			 System.out.println(" "); }
    			 workbook.close();
	}		
	
          	
	
	
	
	
	}

