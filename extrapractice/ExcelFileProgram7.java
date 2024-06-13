package extrapractice;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram7
{
	@Test
	public static void method() throws Exception
	{
		//Take new Excel file with a sheet
		String fp="src/test/resources/excelfiles/book4.xlsx";
		ExcelFileUtility ef=new ExcelFileUtility();
		ef.createXLSXFile(fp);
		Workbook wb=ef.openExcelFile(fp);
		Sheet sh=ef.addSheet(wb, "Sheet1");
		//set a value in a cell with all decorations
		ef.setCellValue(sh, 0, 0, "Hi Students, please wakeup", 14, "Courier New", true, true, 
				IndexedColors.BLUE.getIndex(),IndexedColors.YELLOW.getIndex(), "CENTER");
        //Save excel file
      	ef.saveAndCloseExcel(wb,fp);
	}
}
