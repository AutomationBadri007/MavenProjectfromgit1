package utils;

public class ExcelUtilsMain {

	public static void main(String[] args) 
	{
		String projectpath = System.getProperty("user.dir");
		Excelutils excelpath = new Excelutils(projectpath+"/Exceldata/data.xlsx", "Sheet1");
		excelpath.getRowCount();
		excelpath.getCellDataString(0, 0);
		excelpath.getCellDataNumeric(1, 1);
		

	}

}
