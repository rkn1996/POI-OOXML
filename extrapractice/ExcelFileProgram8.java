package extrapractice;

import java.io.File;
import java.text.SimpleDateFormat;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import utilities.ExcelFileUtility;

public class ExcelFileProgram8 
{
	@Test
	public static void method() throws Exception
	{
		//Create new excel file(.xlsx) with a sheet and headings in 1st row(index=0)
		String fp="src/test/resources/excelfiles/book5.xlsx";
		ExcelFileUtility ef=new ExcelFileUtility();
		ef.createXLSXFile(fp);
		Workbook book=ef.openExcelFile(fp);
		Sheet sh=ef.addSheet(book, "Mysheet");
		ef.setCellValue(sh, 0, 0, "Name",15, "Courier New", true, true, 
				IndexedColors.BLUE.getIndex(),IndexedColors.YELLOW.getIndex(), "CENTER");
		ef.setCellValue(sh, 0, 1, "File/Folder" ,15, "Courier New", true, true,
				IndexedColors.BLUE.getIndex(),IndexedColors.YELLOW.getIndex(), "CENTER");
		ef.setCellValue(sh, 0, 2, "Size",15, "Courier New", true, true, 
				IndexedColors.BLUE.getIndex(),IndexedColors.YELLOW.getIndex(), "CENTER");
		ef.setCellValue(sh, 0, 3, "Last modified",15, "Courier New", true, true, 
				IndexedColors.BLUE.getIndex(),IndexedColors.YELLOW.getIndex(), "CENTER");
		ef.setCellValue(sh, 0, 4, "Hidden",15, "Courier New", true, true, 
				IndexedColors.BLUE.getIndex(),IndexedColors.YELLOW.getIndex(), "CENTER");
		//Copy all files names and other details into excel sheet from 2nd row(index=1)
		File target=new File("E:\\batch265api");
		File[] fl=target.listFiles();
		int rowindex=1; //1st row(index=0) has names for columns in sheet
		for(File f:fl) //take each file or folder from list/collection
		{
			//1. get name of file/folder and then store in 1st column
			ef.setCellValue(sh, rowindex, 0, f.getName());
			sh.autoSizeColumn(0);
			//2. get type and size and then store them into 2nd column and 3rd column
			if(f.isFile())
			{
				ef.setCellValue(sh, rowindex, 1,"file");
				sh.autoSizeColumn(1);
				double b=f.length();
		        double k=(b/1024);
		        ef.setCellValue(sh, rowindex, 2,k+"KB");
				sh.autoSizeColumn(2);
			}
			else
			{
				ef.setCellValue(sh, rowindex, 1,"folder");
				long b=FileUtils.sizeOfDirectory(f);
				double k=(b/1024);
				ef.setCellValue(sh, rowindex, 2,k+"KB");
				sh.autoSizeColumn(2);
			}
			//3. get last modified date and time of file/folder and then store in 4th column
			SimpleDateFormat sdf=new SimpleDateFormat("MMM/dd/yyyy HH:mm:ss");
			ef.setCellValue(sh, rowindex, 3, sdf.format(f.lastModified()));
			sh.autoSizeColumn(3);
			//4. get the file/folder is hidden or not and store in 5th column
			if(f.isHidden())
			{
				ef.setCellValue(sh, rowindex, 4, "Yes");
			}
			else
			{
				ef.setCellValue(sh, rowindex, 4, "No");
			}
			rowindex++; //Mandatory to goto next row in excel sheet
		}
		//Take write permission on that file
		ef.saveAndCloseExcel(book, fp);
	}
}




