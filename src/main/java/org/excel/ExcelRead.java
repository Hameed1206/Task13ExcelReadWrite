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
		
		// Open the Excel file for reading
		FileInputStream fis = new FileInputStream("C:\\Users\\91936\\eclipse-workspace\\ExcelReadandWrite\\Book1.xlsx");
		
		// Create a XSSFWorkbook to represent the Excel workbook
		XSSFWorkbook wbook = new XSSFWorkbook(fis);
		
		// Get the desired sheet (sheet1) from the workbook
		XSSFSheet sheet = wbook.getSheet("sheet1");
		
		 // Loop through each row in the sheet
		for(int i=0; i<=sheet.getLastRowNum(); i++) {
			
			// Get the row at the current index
			XSSFRow row = sheet.getRow(i);
			
			// Loop through each cell in the row
			for (int j = 0; j < row.getLastCellNum(); j++) {
				// Get the cell at the current index
				 XSSFCell cell = row.getCell(j);
				// Get the string value of the cell and print it
				 String value = cell.getStringCellValue();
				 System.out.print(value+ "     ");
			}
			// Move to the next line after printing all cell values in the row
			System.out.println("");
		}
		
		// Close the input stream
        fis.close();
	}
}
