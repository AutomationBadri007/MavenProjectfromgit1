package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelutils 
{
	static String projectpath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public Excelutils(String excelpath, String SheetName)
	{
		try {
		
			 workbook = new XSSFWorkbook("excelpath");
			 sheet = workbook.getSheet("SheetName");
		}
		catch (Exception exp)
		{
			
			exp.printStackTrace();
		}
	}
	
	
	public static void main(String[] args) 
	{
		
		getRowCount();
		getCellDataString(0,0);
		getCellDataNumeric(1,1);
	}
	
	public static void getRowCount()
	{
		
		 
		try {
		
		int rowcount = sheet.getPhysicalNumberOfRows();
		System.out.println("No. of rows:" +rowcount);
		} 
		catch (Exception exp)
		{
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
		
	}
	
	public static void getCellDataString(int rowNum, int columNum)
	{
		try {
			
		 String celldata = sheet.getRow(rowNum).getCell(columNum).getStringCellValue();
		System.out.println(celldata);
		}
		catch (Exception exp)
		{
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
		
	}

	public static void getCellDataNumeric(int rowNum, int columNum)
	{
		try {
			 
		 double celldata = sheet.getRow(rowNum).getCell(columNum).getNumericCellValue();
		System.out.println(celldata);
		}
		catch (Exception exp)
		{
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
		
	}

	
}
