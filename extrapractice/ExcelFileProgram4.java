package extrapractice;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram4
{
	@Test
	public static void method() throws Exception
	{
		//Open an existing file(either ".xls" or ".xlsx")
		ExcelFileUtility ef=new ExcelFileUtility();
		Workbook book=ef.openExcelFile("src\\test\\resources\\excelfiles\\book1.xlsx");
		Sheet sh1=ef.openSheet(book, "Sheet1");
		int nour=ef.getRowsCount(sh1);
		//Create a new sheet for results
		Sheet sh2=ef.addSheet(book, "Sheet2");
		ef.createResultColumn(sh2, 0); //1st column(index=0) in sheet2 is result column
		//Data Driven from 2nd row(index=1) in Sheet1
		for(int i=1;i<nour;i++) //1st row(index=0) has names of columns in Sheet1
		{
			String temp1=ef.getCellValue(sh1, i, 0);
			String temp2=ef.getCellValue(sh1, i, 1);
			int x=Integer.parseInt(temp1);
			int y=Integer.parseInt(temp2);
			int z=x+y;
			String temp3=String.valueOf(z);
			ef.setCellValue(sh2, i, 0, temp3);  //result into 1st column(index=0) in Sheet2	
		}
		//Save excel file into HDD
		ef.saveAndCloseExcel(book,"src\\test\\resources\\excelfiles\\book1.xlsx");
	}
}
