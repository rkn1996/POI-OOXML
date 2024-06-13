package tests;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;

import io.restassured.RestAssured;
import io.restassured.response.Response;

import io.restassured.specification.RequestSpecification;
import utilities.ExcelFileUtility;
import utilities.PropertiesFileUtility;

public class Test2
{
	//Declare global variables and objects
	String pfpath;
	String efpath;
	ExcelFileUtility eu;
	int nour;
	int nouc;
	Workbook book;
	Sheet sh;
	SoftAssert sf;
	
	@BeforeClass
	public void setup() throws Exception
	{
		//Initialize global variables
		pfpath="src\\test\\resources\\propertiesfiles\\config.properties";
		efpath="src\\test\\resources\\excelfiles\\book9_customers_add.xlsx";
		//Define objects
		eu=new ExcelFileUtility();
		sf=new SoftAssert();
		//Open Excel file
		book=eu.openExcelFile(efpath);
		sh=eu.openSheet(book, "Sheet1");
		nour=eu.getRowsCount(sh);
		nouc=eu.getCellscount(sh, 0);
		eu.createResultColumn(sh, nouc);
	}
	
	@Test
	public void apiTest() throws Exception
	{
		//add customers using API
		//Data Driven from 2nd row(index=1)
		for(int i=1;i<nour;i++) //1st row(index=0) has names of columns
		{
			String temp1=eu.getCellValue(sh, i, 0);
			String temp2=eu.getCellValue(sh, i, 1);
			String temp3=eu.getCellValue(sh, i, 2);
			String temp4=eu.getCellValue(sh, i, 3);
			RequestSpecification req=RestAssured.given();
			req.baseUri(PropertiesFileUtility.getValueFromPropertiesFile(pfpath, "baseuri"));
			req.basePath(PropertiesFileUtility.getValueFromPropertiesFile(pfpath, "basepath"));
			req.header("Content-Type","application/x-www-form-urlencoded");
			req.header("Authorization","Bearer "+PropertiesFileUtility.getValueFromPropertiesFile(
					                                                             pfpath, "token"));
			req.formParam("name", temp1);
			req.formParam("email", temp2);
			req.formParam("description", temp3);
			req.formParam("phone", temp4);
			Response res=req.post();
			if(res.getStatusCode()==200 && res.jsonPath().getString("name").equals(temp1)
					&& res.jsonPath().getString("email").equals(temp2) 
					&& res.jsonPath().getString("description").equals(temp3)
					&& res.jsonPath().getString("phone").equals(temp4))
			{
				Reporter.log("Customer "+temp1+" added successfully");
				eu.setCellValue(sh, i, nouc, "Customer "+temp1+" added successfully");
				sf.assertTrue(true);
			}
			else
			{
				Reporter.log("Customer "+temp1+" not added");
				eu.setCellValue(sh, i, nouc, "Customer "+temp1+" not added");
				sf.assertTrue(false);
			}
		}	
	}
	
	@AfterClass
	public void tearDown() throws Exception
	{
		//save and close excel file
		eu.saveAndCloseExcel(book, efpath);
		sf.assertAll();
	}
}
