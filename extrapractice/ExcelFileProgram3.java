package extrapractice;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram3
{
	@Test
	public static void method() throws Exception
	{
		//Open an existing excel file(either ".xls" or ".xlsx")
		ExcelFileUtility ef=new ExcelFileUtility();
		Workbook book=ef.openExcelFile("src\\test\\resources\\excelfiles\\book1.xlsx");
		Sheet sh=ef.openSheet(book,"Sheet1");
		int nour=ef.getRowsCount(sh);
		ef.createResultColumn(sh,2); //index 2 means 3rd column
		//Data Driven from 2nd row(index=1)
		for(int i=1;i<nour;i++) //1st row(index=0) has names of columns
		{
			String temp1=ef.getCellValue(sh, i, 0); //1st column
			String temp2=ef.getCellValue(sh, i, 1); //2nd column
			int x=Integer.parseInt(temp1); //string to integer
			int y=Integer.parseInt(temp2);
			int z=x+y; //perform addition
			String temp3=String.valueOf(z); //integer to string
			ef.setCellValue(sh, i, 2, temp3); //3rd column	
		}
		//Save excel file into HDD
		ef.saveAndCloseExcel(book,"src\\test\\resources\\excelfiles\\book1.xlsx");
	}
}
