package com.excel.task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelWorkTask1 {
    @Test
	public void exceltowrite() throws IOException {
    	File f = new File(System.getProperty("user.dir") +
   			 "\\src\\test\\resources\\Demo Sheet.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Demo page");
		int totalRows = sheet.getPhysicalNumberOfRows();
//    	XSSFRow row = sheet.getRow(5);
//    	row.getCell(2).setCellValue("Yellow Mellow");
//		FileOutputStream out = new FileOutputStream(f);
//		workbook.write(out);
//		workbook.close();
//		out.close();
		    XSSFRow row = sheet.createRow(totalRows);
			row.createCell(0).setCellValue("Java");
			FileOutputStream out = new FileOutputStream(f);
			workbook.write(out);
			workbook.close();
			out.close();
		
		
	
    
	}
}
