package extrapractice;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram6
{
	@Test
	public static void method() throws Exception
	{
		ExcelFileUtility ef=new ExcelFileUtility();
		Workbook wb=ef.openExcelFile("src\\test\\resources\\excelfiles\\Book3.xlsx");
		Sheet sh=ef.openSheet(wb, "Sheet1");
		int nour=ef.getRowsCount(sh);
		int nouc=ef.getCellscount(sh,0);
		//Row sum
		for(int i=0;i<nour;i++)  //row wise
		{
			int rowsum=0;
			for(int j=0;j<nouc;j++) //column wise in every row
			{
				String temp=ef.getCellValue(sh, i, j);
				int x=Integer.parseInt(temp);
				rowsum=rowsum+x;
			}
			ef.setCellValue(sh, i, nouc, String.valueOf(rowsum));
		}
		//column sum
		for(int i=0;i<nouc;i++) //column wise
		{
			int colsum=0;
			for(int j=0;j<nour;j++) //row wise in each column
			{
				String temp=ef.getCellValue(sh, j, i);
				int x=Integer.parseInt(temp);
				colsum=colsum+x;
			}
			ef.setCellValue(sh, nour, i, String.valueOf(colsum));
		}
		//save and close
		ef.saveAndCloseExcel(wb,"src\\test\\resources\\excelfiles\\Book3.xlsx");
	}
}





