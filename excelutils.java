package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelutils {
	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	
		public excelutils(String excelpath, String sheetName) {
			try {
			String projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook (projectpath+"/excel/data.xlsx");
			XSSFSheet sheet =  workbook.getSheet(sheetName);
			}catch(Exception e) {
			e.printStackTrace();
			}
		}
		
	
	public static void main(String[] args) {
		getrowcount();
		getcolcount();
		getcelldataString(0,0);
		getcelldataNumber(2,1);
	}
	public static int getrowcount() {
		int rowCount=0;
		try {
			String projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook (projectpath+"/excel/data.xlsx");
			XSSFSheet sheet =  workbook.getSheet("sheet1");
			
			
			rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("No of rows : "+rowCount);



		} catch(Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
		return rowCount;
	
	}
	public static int getcolcount() {
		
		int colCount=0;
		try {
			String projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook (projectpath+"/excel/data.xlsx");
			XSSFSheet sheet =  workbook.getSheet("sheet1");
			
			
			 colCount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("No of columns : "+colCount);



		} catch(Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
		return colCount;
	}

	public static String getcelldataString(int rowNum, int colNum) {
		
		String celldata= null;
		try {
			String projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook (projectpath+"/excel/data.xlsx");
			XSSFSheet sheet =  workbook.getSheet("sheet1");
			
			 celldata = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println(celldata);

		} catch(Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
		return celldata;
		}

		public static void getcelldataNumber(int rowNum, int colNum) {
			double celldata= 0;
			try {
				String projectpath = System.getProperty("user.dir");
				workbook = new XSSFWorkbook (projectpath+"/excel/data.xlsx");
				XSSFSheet sheet =  workbook.getSheet("sheet1");
				
				 celldata = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
				System.out.println(celldata);

			} catch(Exception exp) {
				System.out.println(exp.getMessage());
				System.out.println(exp.getCause());
				exp.printStackTrace();
				
			}
	}
	}

