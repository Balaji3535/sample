package org.excell;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class ExcelRead {
	public static void main(String[] args) throws IOException {
		
		File loc=new File("C:\\Users\\jeoba\\eclipse-workspace\\kani\\src\\test\\resources\\Book1.xlsx");
		
		
		FileInputStream fi=new FileInputStream(loc);
		
		
		Workbook w=new XSSFWorkbook(fi);
		
		Sheet s = w.getSheet("sheet1");
		
		Row r = s.getRow(1);
		
		Cell cell = r.getCell(2);
		System.out.println(cell);
		
		for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
			Row row = s.getRow(i);
        for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
	
	    Cell cell2 = row.getCell(j);
	    System.out.println(cell2);
	
	
}
		
		}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
