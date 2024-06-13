package extrapractice;

import utilities.ExcelFileUtility;

public class ExcelFileProgram1
{
	public static void main(String[] args) throws Exception
	{
		//Create ".xls" file with a sheet
		ExcelFileUtility ef=new ExcelFileUtility();
		//use single "/" or"\\" in file paths in Java code
		ef.createXLSFile("src/test/resources/excelfiles/Resultbook1.xls");
	}
}
