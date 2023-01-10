package com.example.poi;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiMain {

	public static void main(String[] args) {
		try {
			 InputStream is = new FileInputStream("demo.xlsx");
			 BufferedInputStream bis = new BufferedInputStream(is);
			 Workbook wb = WorkbookFactory.create(bis);
			 Sheet sheet = wb.getSheetAt(0);
			for(int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				Cell cell = row.getCell(0);
				System.out.println("編號：" + cell.getNumericCellValue());
				
				cell = row.getCell(1);
				System.out.println("名字：" + cell.getStringCellValue());
				
				cell = row.getCell(2);
				System.out.println("身分證號：" + cell.getStringCellValue());
				
				cell = row.getCell(3);				 
				DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd hh:mm:ss a");
				LocalDateTime dateTime = cell.getLocalDateTimeCellValue();
				String excelDateTime = dtf.format(dateTime);
				System.out.println("建立日期：" + excelDateTime);
				
				cell = row.getCell(4);
				System.out.println("新增：" + cell.getNumericCellValue());
				System.out.println("============================");
//				if(row.getCell(4) == null) {
//					Row title = sheet.getRow(0);
//					Cell cellTitle = title.createCell(4);
//					cellTitle.setCellType(CellType.STRING);
//					cellTitle.setCellValue("新增");
//					cell = row.createCell(4);
//					cell.setCellType(CellType.NUMERIC);
//					cell.setCellValue(66);
//			    }
//				try (OutputStream fileOut = new FileOutputStream("demo.xlsx")) {
//			        wb.write(fileOut);
//			    }
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
}
