package com.ExcelExample.org;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CompareResponseTimeInTwoDifferentExcelSheet     
{
	public static void main(String[] args) throws IOException {

		FileInputStream file1 = new FileInputStream("./src/test/resources/Data1.xlsx");
		FileInputStream file2 = new FileInputStream("./src/test/resources/Data2.xlsx");
		@SuppressWarnings("resource")
		XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
		@SuppressWarnings("resource")
		XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
		XSSFSheet sheet1 = workbook1.getSheetAt(0);
		XSSFSheet sheet2 = workbook2.getSheetAt(0);
		
		int j=0;
		while (j<=sheet1.getLastRowNum()) {
			if(j==0) {
				j++;
				continue;
			}
			XSSFRow row1 = sheet1.getRow(j);
			if ((row1 == null)) {
				j++;
                continue;
            }
			int k=0;
				XSSFCell cell1 = row1.getCell(k);
				if ((cell1 == null)) {
					j++;
	                continue;
	            }
				switch(cell1.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					String str1 = cell1.getStringCellValue();
					System.out.println(str1);
					System.out.println(j);
					double response2 = passSheet2AndString1(str1,sheet2);
					//System.out.println(response2);
					k+=2;
					cell1 = row1.getCell(k);
					switch(cell1.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						double response1 = cell1.getNumericCellValue();
						//System.out.println(response1);
						if(response1 < response2) {
							System.out.println("Low Response Time in Previous Build");
							System.out.println(response1);
						}
							else {
							System.out.println("Low Response Time in Current Build");
							System.out.println(response2);
						}
						break;
					}
					break;
				}				
			j++;
		}
		file1.close();
		file2.close();
		}
	public static double passSheet2AndString1(String str1,XSSFSheet sheet2) {
		int j=0;
		double res=0;
		while (j<=sheet2.getLastRowNum()) {
			if(j==0) {
				j++;
				continue;
			}
			XSSFRow row2 = sheet2.getRow(j);
			if ((row2 == null)) {
				j++;
                continue;
            }
			int k=0;
				XSSFCell cell2 = row2.getCell(k);
				if ((cell2 == null)) {
					j++;
	                continue;
	            }
				switch(cell2.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					String str2 = cell2.getStringCellValue();
					if(str1.contentEquals(str2)) {
						System.out.println(str2);
						System.out.println(j);
						k+=2;
						cell2 = row2.getCell(k);
						switch(cell2.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							res = cell2.getNumericCellValue();
							break;
						}					
					}
					break;
	            }j++;
		}
		return res;
	}
}