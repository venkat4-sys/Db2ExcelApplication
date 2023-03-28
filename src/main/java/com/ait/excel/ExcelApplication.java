package com.ait.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;



public class ExcelApplication {

	public static void main(String[] args) throws Exception {
		
		String url = "jdbc:mysql://localhost:3306/organization";
        String uname = "root";
        String pwd = "Venkat_147";
        String query = "SELECT * FROM emp";
        
        Connection con = DriverManager.getConnection(url,uname,pwd);
		Statement stmt = con.createStatement();
		ResultSet rs = stmt.executeQuery(query);
		System.out.println("data fetched");
		

		// Create Excel workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("result");

		// Write the column headers to the first row of the worksheet
		Row headerRow = sheet.createRow(0);
		int columnCount = rs.getMetaData().getColumnCount();
		for (int i = 1; i <= columnCount; i++) {
			Cell cell = headerRow.createCell(i - 1);
			cell.setCellValue(rs.getMetaData().getColumnName(i));
		}

		// Write data to sheet
		int rowNum = 1;
		while (rs.next()) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(rs.getInt(1));
			row.createCell(1).setCellValue(rs.getString(2));

			row.createCell(2).setCellValue(rs.getInt(3));
			
			row.createCell(3).setCellValue(rs.getString(4));
			
			
			row.createCell(5).setCellValue(rs.getString(6));
			
		}

		// Save workbook to file
		File file = new File("D:\\data\\result.xlsx");
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();
		System.out.println("EmployeeData.xlsx written successfully!");

	}
		
		
		
      
      
		

	}


