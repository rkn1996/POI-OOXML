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

public class Test5
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
		efpath="src\\test\\resources\\excelfiles\\book12_customers_delete.xlsx";
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
			RequestSpecification req=RestAssured.given();
			req.baseUri(PropertiesFileUtility.getValueFromPropertiesFile(pfpath, "baseuri"));
			req.basePath(PropertiesFileUtility.getValueFromPropertiesFile(
					pfpath, "basepath")+"/"+temp1); //custid as path parameter
			req.header("Content-Type","application/x-www-form-urlencoded");
			req.header("Authorization","Bearer "+PropertiesFileUtility.getValueFromPropertiesFile(
					                                                             pfpath, "token"));
			Response res=req.delete();
			System.out.println(res.getStatusCode());
			if(res.getStatusCode()==200)
			{
				Reporter.log("Customer "+temp1+" deleted successfully");
				eu.setCellValue(sh, i, nouc, "Customer "+temp1+" deleted successfully");
				sf.assertTrue(true);
			}
			else if(res.getStatusCode()==404)
			{
				Reporter.log("Customer "+temp1+" already deleted");
				eu.setCellValue(sh, i, nouc, "Customer "+temp1+" already deleted");
				sf.assertTrue(true);
			}
			else
			{
				Reporter.log("Customer "+temp1+" not updated");
				eu.setCellValue(sh, i, nouc, "Customer "+temp1+" not deleted");
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
