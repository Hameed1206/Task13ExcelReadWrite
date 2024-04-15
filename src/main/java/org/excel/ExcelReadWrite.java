package org.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelReadWrite {

	public static void main(String[] args) throws IOException {
		
		//  Create a FileInputStream to read the Excel file
		FileInputStream fis = new FileInputStream("C:\\Users\\91936\\eclipse-workspace\\ExcelReadandWrite\\Book1.xlsx");
		
		// Create a XSSFWorkbook to represent the Excel workbook
		XSSFWorkbook wbook = new XSSFWorkbook(fis);
		
		// Get the desired sheet (sheet1) from the workbook
		XSSFSheet sheet = wbook.getSheet("sheet1");
		
		// Create a row at index 0 in the sheet
		XSSFRow row = sheet.createRow(0);
		
		// Create cells at index 0, 1 and 2 in the row and set values
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Name");
		
		cell = row.createCell(1);
		cell.setCellValue("Age");
		
		cell = row.createCell(2);
		cell.setCellValue("Email");
		
		// Create a row at index 1 in the sheet
		row = sheet.createRow(1);
		
		// Create cells at index 0, 1 and 2 in the row and set values
		cell = row.createCell(0);
		cell.setCellValue("John Doe");
		
		cell = row.createCell(1);
		cell.setCellValue("30");
		
		cell = row.createCell(2);
		cell.setCellValue("john@test.com");
		
		// Create a row at index 2 in the sheet
        row = sheet.createRow(2);
		
        // Create cells at index 0, 1 and 2 in the row and set values
		cell = row.createCell(0);
		cell.setCellValue("Jane Doe");
		
		cell = row.createCell(1);
		cell.setCellValue("28");
		
		cell = row.createCell(2);
		cell.setCellValue("john@test.com");
		
		// Create a row at index 3 in the sheet
        row = sheet.createRow(3);
		
        // Create cells at index 0, 1 and 2 in the row and set values
		cell = row.createCell(0);
		cell.setCellValue("Bob Smith");
		
		cell = row.createCell(1);
		cell.setCellValue("35");
		
		cell = row.createCell(2);
		cell.setCellValue("jacky@test.com");
		
		// Create a row at index 4 in the sheet
        row = sheet.createRow(4);
		
        // Create cells at index 0, 1 and 2 in the row and set values
		cell = row.createCell(0);
		cell.setCellValue("Swapnil");
		
		cell = row.createCell(1);
		cell.setCellValue("37");
		
		cell = row.createCell(2);
		cell.setCellValue("swapnil@test.com");
		
		// Create a FileOutputStream to write to the Excel file
		FileOutputStream fos = new FileOutputStream("C:\\\\Users\\\\91936\\\\eclipse-workspace\\\\ExcelReadandWrite\\\\Book1.xlsx");
		// To write the workbook
		wbook.write(fos);
		
		// Close the input , output streams and worbook.
		fis.close();
		fos.close();
		wbook.close();

		 // Output a message indicating successful completion
        System.out.println("Excel file updated successfully.");
	}

}
