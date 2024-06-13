package extrapractice;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;

import utilities.BrowserWindowUtility;
import utilities.ExcelFileUtility;
import utilities.JavaScriptUtility;
import utilities.WebSiteUtility;

public class ExcelFileProgram10
{
	//Declare global objects and variables
	WebSiteUtility ws;
	WebDriver driver;
	FluentWait<WebDriver> wait;
	BrowserWindowUtility bu;
	SoftAssert sa;
	JavaScriptUtility ju;
	ExcelFileUtility ef;
	String exceFilePath;
	Workbook book;
	Sheet sh;
	int nour;
		
	@BeforeClass
	public void testSetup() throws Exception
	{
		//Create objects to utility classes
		ws=new WebSiteUtility();
		bu=new BrowserWindowUtility();
		ju=new JavaScriptUtility();
		ef=new ExcelFileUtility();
		sa=new SoftAssert();
		//open excel file
		exceFilePath="src/test/resources/excelfiles/book7_redbus_testdata.xlsx";
		book=ef.openExcelFile(exceFilePath);
		sh=ef.openSheet(book, "Sheet1");
		nour=ef.getRowsCount(sh);
	}
	
	@Test(priority=1)
	public void redBusLaunch() throws Exception
	{
		//Open browser, Define wait object, and launch site
		ChromeOptions options = new ChromeOptions();
	    options.addArguments("--disable-notifications");
	    driver=new ChromeDriver(options);
		bu.browserMaximize(driver);
		wait=ws.defineExplicitWait(driver, 10, 1000);
		ws.launchSite(driver,"https://www.redbus.in/");
		if(driver.getTitle().contains("Book Bus Tickets Online"))
		{
			sa.assertTrue(true);
		}
		else
		{
			sa.fail();
		}
		if(driver.getCurrentUrl().startsWith("https"))
		{
			sa.assertTrue(true);
		}
		else
		{
			sa.fail();
		}
	}
	
	@Test(priority=2, dependsOnMethods="redBusLaunch")
	public void redBusMultiSearch() throws Exception
	{
		//Data Driven from 2nd row(index=1)
		for(int i=1; i<nour; i++) //1st row(index=0) has names of columns
		{
			String x=ef.getCellValue(sh, i, 0);
			String y=ef.getCellValue(sh, i, 1);
			String z=ef.getCellValue(sh, i, 2);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("src"))).clear();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("src")))
		                                                                        .sendKeys(x);
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
				                               "//ul[contains(@class,'sc')]/li[1]"))).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("dest"))).clear();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("dest")))
																		         .sendKeys(y);
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
											   "//ul[contains(@class,'sc')]/li[1]"))).click();
			wait.until(ExpectedConditions.elementToBeClickable(By.id("onwardCal"))).click();
			//back to current month if "previous" icon exists
			try
			{
				WebElement previous=wait.until(ExpectedConditions.visibilityOfElementLocated(
					By.xpath(
				 "(//div[@id='onwardCal']//div[contains(@class,'DayNavigator__IconBlock')])[1]")));
				while(true)
				{
					previous.click();
					Thread.sleep(2000);
				}
			}
			catch(Exception ex)
			{	
			}
			//goto target month and year
			String pieces[]=z.split("-");
			String day=pieces[0];
			String month=pieces[1];
			String year=pieces[2];
			WebElement text=wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
				"(//div[@id='onwardCal']//div[contains(@class,'DayNavigator__IconBlock')])[2]")));
			WebElement next=wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
				"(//div[@id='onwardCal']//div[contains(@class,'DayNavigator__IconBlock')])[3]")));
			while(true)
			{
				if(text.getText().toLowerCase().contains(month.toLowerCase()) && 
					text.getText().contains(year))
				{
					break;
				}
				else
				{
					next.click();
					Thread.sleep(2000);
				}
			}
			//goto target day and click on search
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
				                     "//div[@id='onwardCal']//span[text()='"+day+"']"))).click();
			wait.until(ExpectedConditions.elementToBeClickable(By.id("search_button"))).click();
			//close extra banners if exists
			try {
				wait.until(ExpectedConditions.elementToBeClickable(By.className("gotIt"))).click();
			}
			catch(Exception ex){	
			}
			try {
				wait.until(ExpectedConditions.elementToBeClickable(By.className("close-primo")))
			                                                                            .click();
			}
			catch(Exception ex){	
			}
			// Set a variable to keep track of the last height
			long lastHeight=(long) ((JavascriptExecutor) driver).executeScript(
        		                                             "return document.body.scrollHeight");
			// Loop until all elements(buses information) are loaded
			while(true) 
			{
				// Scroll down the page
				ju.scrollPageDownByJS(driver);
				// Wait for some time to allow content to load
				Thread.sleep(2000); // Adjust this delay as needed
				// Calculate the new height
				long newHeight = (long) ((JavascriptExecutor) driver).executeScript(
            		                                         "return document.body.scrollHeight");
				// If the height has not changed, we've reached the bottom of the page
				if(newHeight == lastHeight) 
				{
					break;
				}
				// Update the last height
				lastHeight = newHeight;
			} 
			List<WebElement> travelbuses=driver.findElements(By.xpath(
					                                        "//ul[@class='bus-items']/div"));
			// find count of buses, maximum fare and minimum fare.
			ef.setCellValue(sh, i, 3, String.valueOf(travelbuses.size()));
			int maxfare=0; //assume small to change to maximum
			int minfare=10000; //assume big to change to minimum
			String minbus="";
			String maxbus="";
			for(WebElement travelbus : travelbuses) {
				String travelname=travelbus.findElement(
    		        By.xpath("descendant::div[starts-with(@class,'travel')]")).getText();
				String dl=travelbus.findElement(
    		        By.xpath("descendant::div[starts-with(@class,'dp-loc')]")).getText();
				String dt=travelbus.findElement(
    		        By.xpath("descendant::div[starts-with(@class,'dp-time')]")).getText();
				String al=travelbus.findElement(
    		        By.xpath("descendant::div[starts-with(@class,'bp-loc')]")).getText();
				String at=travelbus.findElement(
    		        By.xpath("descendant::div[starts-with(@class,'bp-time')]")).getText();
				String fare=travelbus.findElement(
    		        By.xpath("descendant::div[starts-with(@class,'fare')]")).getText();
				fare=fare.replace("INR","");
				fare=fare.trim();
				int farevalue=Integer.parseInt(fare);
				if(farevalue<minfare)
				{
					minfare=farevalue;
					minbus=travelname+" "+dl+" "+dt+" "+al+" "+at+" "+fare;
				}
				if(farevalue>maxfare)
				{
					maxfare=farevalue;
					maxbus=travelname+" "+dl+" "+dt+" "+al+" "+at+" "+fare;
				}
			}
			ef.setCellValue(sh, i, 4, minbus);
			ef.setCellValue(sh, i, 5, maxbus);
			//Re-launch site to back to home page
			driver.navigate().to("https://www.redbus.in/");
		}
	}
	
	@Test(priority=3, dependsOnMethods="redBusLaunch")
	public void redBusClose()
	{
        //close site
        ws.closeSite(driver); 
	}
	
	@AfterClass
	public void tearDown() throws Exception
	{
		//save and close excel file
		ef.saveAndCloseExcel(book,exceFilePath);
		sa.assertAll();
	}
}
