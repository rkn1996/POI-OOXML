package tests;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.restassured.RestAssured;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import utilities.BrowserWindowUtility;
import utilities.PropertiesFileUtility;
import utilities.WebSiteUtility;

public class Test1
{
	//Global Variables and object declaration
	String pfpath;
	WebSiteUtility ws;
	BrowserWindowUtility bu;
	WebDriver driver;
	FluentWait<WebDriver> wait;
	int uivalue;
	int apivalue;
	
	@BeforeClass
	public void testSetup()
	{
		//initialize variables
		pfpath="src\\test\\resources\\propertiesfiles\\config.properties";
		//create objects
		ws=new WebSiteUtility();
		bu=new BrowserWindowUtility();
	}
		
	@Test(priority=1)
	public void apiTest() throws Exception
	{
		RequestSpecification req=RestAssured.given();
		req.baseUri(PropertiesFileUtility.getValueFromPropertiesFile(pfpath,"baseuri"));
		req.basePath(PropertiesFileUtility.getValueFromPropertiesFile(pfpath,"basepath"));
		req.header("Authorization","Bearer "+PropertiesFileUtility.getValueFromPropertiesFile(
				                                                              pfpath,"token"));
	    Response res=req.get();
	    apivalue=res.jsonPath().getInt("data.size()"); //Gpath expression
	    System.out.println(apivalue);
	}
	
	@Test(priority=2)   //giving security error when we access UI via automation
	public void uiTest() throws Exception
	{
		//open browser
		driver=ws.openBrowser("chrome");
		//launch site
		String url=PropertiesFileUtility.getValueFromPropertiesFile(pfpath,"url");
		ws.launchSite(driver, url);
		//maximize browser window
		bu.browserMaximize(driver);
		//Define wait
		int m=Integer.parseInt(PropertiesFileUtility.getValueFromPropertiesFile(
				                                                     pfpath,"maxwaittimeinsecs"));
		int i=Integer.parseInt(PropertiesFileUtility.getValueFromPropertiesFile(
				                                                   pfpath,"intervaltimeinmsecs"));
		wait=ws.defineExplicitWait(driver, m, i);
		//do login
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("email")))
		    .sendKeys(PropertiesFileUtility.getValueFromPropertiesFile(pfpath,"username"));
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("old-password")))
		    .sendKeys(PropertiesFileUtility.getValueFromPropertiesFile(pfpath,"password"));
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@type='submit']")))
		    .click();
		//click on customers
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
				        "//a[@data-testid='primary-nav-item-link-customers']"))).click();
		//get count of customers
		uivalue=wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(
			"//div[contains(@data-moduleid,'listview-customers')]//table/tbody/tr"))).size();
		//do logout
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@aria-label='Settings']")))
																						.click();
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
			                                          "//*[contains(text(),'Sign out')]"))).click();
		//close browser
		ws.closeSite(driver);
		System.out.println(uivalue);
	}
		
	@Test(priority=3)
	public void e2eTest()
	{
		if(uivalue==apivalue)
		{
			Reporter.log("E2E test passed");
		}
		else
		{
			Reporter.log("E2E test failed");
			Assert.fail(); //hard assertion
		}
	}
}
