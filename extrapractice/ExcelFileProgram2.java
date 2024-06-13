package extrapractice;

import java.util.Scanner;

import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram2
{
	@Test
	public static void method() throws Exception
	{
		Scanner sc=new Scanner(System.in);
		System.out.println("How many .xlsx files you want to create?");
		int n=sc.nextInt();
		sc.close();
		for(int i=1;i<=n; i++)
		{
			//Create ".xlsx" file with a sheet
			ExcelFileUtility ef=new ExcelFileUtility();
			ef.createXLSXFile("src/test/resources/excelfiles/Resultbook"+i+".xlsx");
		}
	}
}
