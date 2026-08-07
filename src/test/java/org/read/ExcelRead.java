package org.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.imageio.stream.FileImageInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.binary.XSSFBCommentsTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.bytebuddy.asm.Advice.OffsetMapping.ForOrigin;

public class ExcelRead {
    public static void main(String[] args) throws IOException {
	
    	File f =new File("src\\test\\resources\\TestData\\Read.xlsx");
    	FileInputStream fis=new FileInputStream(f);
    	Workbook book = new XSSFWorkbook(fis);
    	Sheet sheet = book.getSheet("Sheet1");
    	for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
    		Row row = sheet.getRow(i);
    	for (int j =0;j<row.getPhysicalNumberOfCells();j++) {
    		Cell cell = row.getCell(j);
    		CellType cells = cell.getCellType();
    		switch(cells) {
    		case STRING:
    			String scv = cell.getStringCellValue();
    			System.out.println(scv +"\t");
    			break;
    		case NUMERIC:
    			if(DateUtil.isCellDateFormatted(cell)) {
    				Date date = cell.getDateCellValue();
    				SimpleDateFormat sdf=new SimpleDateFormat("MM-dd-yyyy");
    				String date1 = sdf.format(date);
    				System.out.println(date1+"\t");
    				}
    			else {
    				double numericValue = cell.getNumericCellValue();
    				long l=(long)numericValue;
    				System.out.println(l+"\t");
    			}
    			break;
    		case BOOLEAN:
    			boolean bcv = cell.getBooleanCellValue();
    			System.out.println(bcv+"\t");
    			break;
    		case BLANK:
    			System.out.println("BLANK CELL"+"\t");
    			break;
    		default:
    			System.out.println("Invalid Cell Type"+"\t");
    			break;
    			
    		}}
    	System.out.println();
    	}}}
