package extrapractice;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.restassured.RestAssured;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import utilities.ExcelFileUtility;
import utilities.WebSiteUtility;

public class ExcelFileProgram11
{
	//declare global variables and objects
	String efpath;
	int apivalue;
	int uivalue;
	int dbvalue;
	WebDriver driver;
	FluentWait<WebDriver> wait;
	ExcelFileUtility eu;
	Workbook book;
	Sheet sh;
	WebSiteUtility ws;
	
	@BeforeClass
	public void testSetup() throws Exception
	{
		//initialize global variables
		efpath="src\\test\\resources\\excelfiles\\book8_cricbuz_testdata.xlsx";
		//create objects
		eu=new ExcelFileUtility();
		ws=new WebSiteUtility();
		//open excel file
		book=eu.openExcelFile(efpath);
		sh=eu.openSheet(book, "Sheet1");
	}
	
	
	@Test
	public void e2eTest() throws Exception
	{
		//Data driven testing from 2nd row(index=1)
		for(int i=1;i<eu.getRowsCount(sh);i++) //1st row(index=0) has names of columns
		{
			String url=eu.getCellValue(sh, i, 0); //1st column
			String matchid=eu.getCellValue(sh, i, 1); //2nd column
			String innings=eu.getCellValue(sh, i, 2); //3rd column
			String batsman=eu.getCellValue(sh, i, 3); //4th column
			int in=Integer.parseInt(innings);
			int b=Integer.parseInt(batsman);
			//1. contact API(RA-Java)
			RequestSpecification req=RestAssured.given();
			req.baseUri("https://cricbuzz-cricket.p.rapidapi.com");
			req.basePath("mcenter/v1/"+matchid+"/hscard");
			req.header("X-RapidAPI-Key","71f3c982bemshe28bcaeabbcb163p1d75dbjsn724f99e47c76");
			req.header("X-RapidAPI-Host","cricbuzz-cricket.p.rapidapi.com");
			Response res=req.get();
			apivalue=res.jsonPath().getInt(
	    		"scoreCard["+(in-1)+"].batTeamDetails.batsmenData.bat_"+b+".runs");
			eu.setCellValue(sh, i, 5, String.valueOf(apivalue));
			//2. contact UI(SWD-Java)
			driver=ws.openBrowser("chrome");
			ws.launchSite(driver, url);
			wait=new FluentWait<WebDriver>(driver);
			WebElement target=wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
 "(//div[@id='innings_"+innings+"']//div[contains(@class,'cb-scrd-itms')])["+batsman+"]/div[3]")));
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();",target); 
			Thread.sleep(3000);
			((JavascriptExecutor) driver).executeScript(
					"arguments[0].style.border='2px solid red';",target);
			Thread.sleep(3000);
			String temp=target.getText();
			uivalue=Integer.parseInt(temp);
			eu.setCellValue(sh, i, 4, String.valueOf(uivalue));
			driver.quit();
			//3. contact DB
		}
	}
	
	@AfterClass
	public void tearDown() throws Exception
	{
		//save and close excel file
		eu.saveAndCloseExcel(book,efpath);
	}
}
