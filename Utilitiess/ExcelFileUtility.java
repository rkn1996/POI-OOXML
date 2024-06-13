package utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//To access Excel file
public class ExcelFileUtility 
{
	public void createXLSFile(String filepath) throws Exception
	{
		//Create ".xls" file with a sheet
		File f=new File(filepath);
		FileOutputStream fo=new FileOutputStream(f); //write mode
		HSSFWorkbook hwb=new HSSFWorkbook(); 
		hwb.createSheet();
		hwb.write(fo); //save xls file
		hwb.close();
	}
	
	public void createXLSXFile(String filepath) throws Exception
	{
		//Create ".xlsx" file with a sheet
		File f=new File(filepath);
		FileOutputStream fo=new FileOutputStream(f); //write mode
		XSSFWorkbook xwb=new XSSFWorkbook(); 
		xwb.createSheet();
		xwb.write(fo); //save xlsx file
		xwb.close();
	}
	
	public Workbook openExcelFile(String filepath) throws Exception
	{
		//either ".xls" or ".xlsx"
		File f=new File(filepath);
		FileInputStream fi=new FileInputStream(f); //READ mode
		Workbook wb=WorkbookFactory.create(fi);
		return(wb);
	}
	
	public Sheet openSheet(Workbook wb, Object input) 
	{
		Sheet sht;
        if(input instanceof Integer){
            int intValue = (int) input;
            sht=wb.getSheetAt(intValue);
            return(sht);
        } 
        else if (input instanceof String){
            String stringValue = (String) input;
            sht=wb.getSheet(stringValue);
            return(sht);
        } 
        else{
        	throw new NullPointerException();
        }
	}
	
	public Sheet addSheet(Workbook wb, String sheetName)
	{
		Sheet sht=wb.createSheet(sheetName);
		return(sht);
	}
	
	public int getRowsCount(Sheet sht)
	{
		int nour=sht.getPhysicalNumberOfRows();
		return(nour);
	}
	
	public int getCellscount(Sheet sht, int rowindex)
	{
		int nouc=sht.getRow(rowindex).getLastCellNum(); 
		return(nouc);
	}
	
	public void createResultColumn(Sheet sht, int cellindex)
	{
		SimpleDateFormat sf=new SimpleDateFormat("dd-MMM-yyyy-hh-mm-ss");
		Date dt=new Date();
		//create results column in first row by default
		try {
			sht.getRow(0).createCell(cellindex).setCellValue("Test Results on "+sf.format(dt));
		}
		catch(Exception ex)
		{
			sht.createRow(0).createCell(cellindex).setCellValue("Test Results on "+sf.format(dt));
		}
		Font font=sht.getWorkbook().createFont();
		font.setFontHeightInPoints((short)15);
		font.setBold(true);
		font.setItalic(true);
		CellStyle style=sht.getWorkbook().createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.valueOf("CENTER"));
		sht.getRow(0).getCell(cellindex).setCellStyle(style);
		sht.autoSizeColumn(cellindex); //auto-fit
	}
	
	public String getCellValue(Sheet sht, int rowindex, int cellindex)
	{
		DataFormatter df=new DataFormatter(); //get any type of data in cell as String only
		String value=df.formatCellValue(sht.getRow(rowindex).getCell(cellindex));
		return(value);
	}
	
	public void setCellValue(Sheet sht, int rowindex, int cellindex, String value) throws Exception
	{
		try
		{
			//If row is used row
			sht.getRow(rowindex).createCell(cellindex).setCellValue(value);
			sht.autoSizeColumn(cellindex); 
		}
		catch(NullPointerException ex)
		{
			//If row is unused or using first time
			sht.createRow(rowindex).createCell(cellindex).setCellValue(value);
			sht.autoSizeColumn(cellindex); 
		}
	}
	
	public void setCellValue(Sheet sht, int rowindex, int cellindex, String value, int fontSize,
	  String fontName, boolean bold, boolean italic, short fontColor, short bgColor, String align)
	{
		Font font=sht.getWorkbook().createFont();
		font.setFontName(fontName);
		font.setFontHeightInPoints((short)fontSize);
		font.setBold(bold);
		font.setItalic(italic);
		font.setColor(fontColor);
		CellStyle style=sht.getWorkbook().createCellStyle();
		style.setFont(font);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND); // Set the fill pattern
		style.setFillForegroundColor(bgColor); // Set the fill foreground color
		style.setAlignment(HorizontalAlignment.valueOf(align));
		try
		{
			//If row is used row
			sht.getRow(rowindex).createCell(cellindex).setCellValue(value);
		}
		catch(Exception ex)
		{
			//If row is unused or using first time
			sht.createRow(rowindex).createCell(cellindex).setCellValue(value);
		}
		sht.getRow(rowindex).getCell(cellindex).setCellStyle(style);
		sht.autoSizeColumn(cellindex);
	}
	
	public void saveAndCloseExcel(Workbook wb, String filepath) throws Exception
	{
		FileOutputStream fo=new FileOutputStream(filepath);
		wb.write(fo); //save
	}
}
