package tests;

import java.util.List;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.restassured.RestAssured;
import io.restassured.response.Response;

import io.restassured.specification.RequestSpecification;
import utilities.ExcelFileUtility;
import utilities.PropertiesFileUtility;

public class Test3
{
	//Declare global variables and objects
	String pfpath;
	String efpath;
	ExcelFileUtility eu;
	Workbook book;
	Sheet sh;
	
	@BeforeClass
	public void setup() throws Exception
	{
		//Initialize global variables
		pfpath="src\\test\\resources\\propertiesfiles\\config.properties";
		efpath="src\\test\\resources\\excelfiles\\book10_all_customers.xlsx";
		//Define objects
		eu=new ExcelFileUtility();
		//create and Open Excel file
		eu.createXLSXFile(efpath);
		book=eu.openExcelFile(efpath);
		sh=eu.addSheet(book, "Sheet1");
		eu.setCellValue(sh,0,0,"CustID",15,"Calibri",true,true, IndexedColors.BLUE.getIndex(),
				                                 IndexedColors.YELLOW.getIndex(), "CENTER");
	}
	
	@Test
	public void apiTest() throws Exception
	{
		//Get all customer Ids using API
		RequestSpecification req=RestAssured.given();
		req.baseUri(PropertiesFileUtility.getValueFromPropertiesFile(pfpath, "baseuri"));
		req.basePath(PropertiesFileUtility.getValueFromPropertiesFile(pfpath, "basepath"));
		req.header("Authorization","Bearer "+PropertiesFileUtility.getValueFromPropertiesFile(
					                                                             pfpath, "token"));
		Response res=req.get();
		List<String> ids=res.jsonPath().getList("data.id");
		int rowindex=1;
		for(String id:ids)
		{
			eu.setCellValue(sh,rowindex,0,id);
			rowindex++;
		}
	}
	
	@AfterClass
	public void tearDown() throws Exception
	{
		//save and close excel file
		eu.saveAndCloseExcel(book, efpath);
	}
}
