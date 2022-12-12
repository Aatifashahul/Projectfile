package com.excel.task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelWorkTask2 {
	@Test
	public void exceltocreate() throws IOException {
		File f = new File(System.getProperty("user.dir") +
				"\\src\\test\\resources\\Demo Sheet.xlsx");
				FileInputStream input = new FileInputStream(f);
				XSSFWorkbook workbook = new XSSFWorkbook(input);
				XSSFSheet sheet = workbook.getSheet("Demo page");
				int totalrow = sheet.getPhysicalNumberOfRows();
				XSSFRow row = sheet.createRow(totalrow);
				row.createCell(0).setCellValue("Java");
				FileOutputStream out = new FileOutputStream(f);
				workbook.write(out);
				out.close();
				workbook.close();
	
    
	}
}

