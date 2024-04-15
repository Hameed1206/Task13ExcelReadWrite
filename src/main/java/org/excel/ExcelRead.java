package org.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream("C:\\Users\\91936\\eclipse-workspace\\ExcelReadandWrite\\Book1.xlsx");
		XSSFWorkbook wbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = wbook.getSheet("sheet1");
		
		for(int i=0; i<=sheet.getLastRowNum(); i++) {
			
			XSSFRow row = sheet.getRow(i);
			
			for (int j = 0; j < row.getLastCellNum(); j++) {
				 XSSFCell cell = row.getCell(j);
				 String value = cell.getStringCellValue();
				 System.out.print(value+ "     ");
			}
			System.out.println("");
		}

	}

}
