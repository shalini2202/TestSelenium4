package com.TestSelenium4;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException  {
		// TODO Auto-generated method stub
		
		System.out.println("om");
		
		String filePath = "..//TestSelenium4//src/com//TestSelenium4//TestSelenium4.xlsx";
		FileInputStream inputStream = new FileInputStream(filePath);
		Workbook wb = new XSSFWorkbook(inputStream);
		 

		    //Read excel sheet by sheet name    

		    Sheet sheet = wb.getSheet("Sheet1");
		    
		    sheet.getRow(0).createCell(0).setCellValue("Shalini");
		    FileOutputStream fout=new FileOutputStream(filePath);
		    wb.write(fout);
		    fout.close();
	}

}
