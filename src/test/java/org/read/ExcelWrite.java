package org.read;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
		public static void main(String[] args) throws IOException {
			
			File f = new File("src\\test\\resources\\TestData\\write.xlsx");
			
			Workbook book = new XSSFWorkbook();
			
			Sheet sheet = book.createSheet("customer details");
			
			Row row1 = sheet.createRow(0);
			Row row2 = sheet.createRow(1);
			Row row3 = sheet.createRow(2);
			
			
			Cell cell1 = row1.createCell(0);
			Cell cell2 = row1.createCell(1);
			Cell cell3 = row1.createCell(2);
			Cell cell4 = row1.createCell(3);
			
			Cell cell5 = row2.createCell(0);
			Cell cell6 = row2.createCell(1);
			Cell cell7 = row2.createCell(2);
			Cell cell8 = row2.createCell(3);
			
			Cell cell9  = row3.createCell(0);
			Cell cell10 = row3.createCell(1);
			Cell cell11 = row3.createCell(2);
			Cell cell12 = row3.createCell(3);
			
			cell1.setCellValue("s.no");
			cell2.setCellValue("Name");
			cell3.setCellValue("phno");
			cell4.setCellValue("mail");
			
			cell5.setCellValue(1);
			cell6.setCellValue("Esakkiammal");
			cell7.setCellValue("7904452160");
			cell8.setCellValue("Esakki0422@gmail.com");
			
			cell9.setCellValue(2);
			cell10.setCellValue("Gopinathan");
			cell11.setCellValue("9029689758");
			cell12.setCellValue("Gopi1003@gmail.com");
			
			FileOutputStream fos=new FileOutputStream(f);
			book.write(fos);
			
			}}
