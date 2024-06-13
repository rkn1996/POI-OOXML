package extrapractice;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram5
{
	@Test
	public static void method() throws Exception
	{
		//Open an existing file, get existing Sheet1, get used rows count from Sheet1
		ExcelFileUtility ef=new ExcelFileUtility();
		Workbook book1=ef.openExcelFile("src\\test\\resources\\excelfiles\\book1.xlsx");
		Sheet sh1=ef.openSheet(book1,"Sheet1");
		int nour=ef.getRowsCount(sh1);
		//Create and open a new Excel file with a sheet and result column
		ef.createXLSXFile("src\\test\\resources\\excelfiles\\book2.xlsx");
		Workbook book2=ef.openExcelFile("src\\test\\resources\\excelfiles\\book2.xlsx");
		Sheet sh2=ef.addSheet(book2,"ResultSheet");
		ef.createResultColumn(sh2, 0); //1st column in sheet1 is result column
		//Data Driven from 2nd row(index=1) because 1st row(index=0) has names of columns
		for(int i=1;i<nour;i++)
		{
			String temp1=ef.getCellValue(sh1, i, 0);
			String temp2=ef.getCellValue(sh1, i, 1);
			int x=Integer.parseInt(temp1);
			int y=Integer.parseInt(temp2);
			int z=x+y;
			String temp3=String.valueOf(z);
			ef.setCellValue(sh2, i, 0, temp3);		
		}
		//Save 2nd excel file into HDD
		ef.saveAndCloseExcel(book2, "src\\test\\resources\\excelfiles\\book2.xlsx");
	}
}
