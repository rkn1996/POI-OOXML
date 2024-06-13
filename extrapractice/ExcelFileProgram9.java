package extrapractice;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;

import utilities.BrowserWindowUtility;
import utilities.ExcelFileUtility;
import utilities.RandomUtility;
import utilities.WebSiteUtility;

public class ExcelFileProgram9 
{
	WebDriver driver;
	FluentWait<WebDriver> wait;
	WebSiteUtility ws;
	BrowserWindowUtility bu;
	String exceFilePath;
	ExcelFileUtility ef;
	Workbook book;
	Sheet sh;
	SoftAssert sf;
	
	@BeforeClass
	public void testSetup()
	{
		ef=new ExcelFileUtility();
		ws=new WebSiteUtility();
		bu=new BrowserWindowUtility();
		sf=new SoftAssert();
	}
		
	@Test(priority = 1)
	public void method1() throws Exception
	{
		//take data from excel file and do registration
		exceFilePath="src/test/resources/excelfiles/book6_aidaform_testdata.xlsx";
		book=ef.openExcelFile(exceFilePath);
		sh=ef.openSheet(book, "Sheet1");
		int nour=ef.getRowsCount(sh);
		int nouc=ef.getCellscount(sh,0);
		//create a result column
		ef.createResultColumn(sh,nouc); //next to last column
		//Data Driven from 2nd row(index=1)
		for(int i=1;i<nour;i++) //1st row(index=0) has names of columns
		{
			//open browser and launch site
			driver=ws.openBrowser("chrome");
			bu.browserMaximize(driver);
			ws.launchSite(driver,"https://my.aidaform.com/signup");
			//define wait
			wait=ws.defineExplicitWait(driver, 10, 100);
			//get data from excel file
			String x=ef.getCellValue(sh, i, 0);
			String y=ef.getCellValue(sh, i, 1);
			String z=ef.getCellValue(sh, i, 2);
			String w=ef.getCellValue(sh, i, 3);
			String criteria=ef.getCellValue(sh, i, 4);
			//find elements and enter data
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("nickname")))
			                                                                      .sendKeys(x);
			if(y.equalsIgnoreCase("yes"))
			{
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("email")))
				                                 .sendKeys(RandomUtility.generateRandomEmail());
			}
			else if(y.equalsIgnoreCase("no"))
			{
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("email")))
				                                                                  .sendKeys("");
			}
			else
			{
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("email")))
				                                                                    .sendKeys(y);
			}
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("password")))
			                                                                       .sendKeys(z);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("confirm")))
			                                                                       .sendKeys(w);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
					                                        "(//*[name()='svg'])[4]"))).click();
		
			//submit data
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
					                     "//button[text()='Create My Free Account']"))).click();
			Thread.sleep(5000); //wait for 5 seconds before going to test outcome
			//test outcome
			try{
				if(criteria.equalsIgnoreCase("all_valid") && driver.findElement(
					    By.xpath("//h3[text()=\"Let's Confirm Your Email!\"]")).isDisplayed())
				{
				Reporter.log("Test case passed for valid data");
				ef.setCellValue(sh, i, nouc, "Test case passed for valid data",14,"Courier New",
					true, true,IndexedColors.GREEN.getIndex(),IndexedColors.YELLOW.getIndex(),
					"CENTER");
				}
				else if(criteria.equalsIgnoreCase("uname_blank") && driver.findElement(By.xpath(
					                "//div[text()='Please enter a username']")).isDisplayed())
				{
				Reporter.log("Test case passed for blank username");
				ef.setCellValue(sh,i,nouc,"Test case passed for blank username",14,"Courier New",
						true, true,IndexedColors.GREEN.getIndex(),IndexedColors.YELLOW.getIndex(),
						"CENTER");
				}
				else if(criteria.equalsIgnoreCase("email_blank") && driver.findElement(By.xpath(
					"//div[text()='Please enter your email address']")).isDisplayed())
				{
				Reporter.log("Test case passed for blank email");
				ef.setCellValue(sh, i, nouc,"Test case passed for blank email",14,"Courier New",
						true, true,IndexedColors.GREEN.getIndex(),IndexedColors.YELLOW.getIndex(),
						"CENTER");
				}
				else if(criteria.equalsIgnoreCase("pwd_blank") && driver.findElement(
					By.xpath("//div[text()='Please enter a password']")).isDisplayed())
				{
				Reporter.log("Test case passed for blank password");
				ef.setCellValue(sh,i,nouc,"Test case passed for blank password",14,"Courier New",
						true, true,IndexedColors.GREEN.getIndex(),IndexedColors.YELLOW.getIndex(),
						"CENTER");
				}
				else if(criteria.equalsIgnoreCase("confirmpwd_blank") && driver.findElement(
					By.xpath("//div[text()='Please confirm your password']")).isDisplayed())
				{
				Reporter.log("Test case passed for blank confirm password");
				ef.setCellValue(sh,i,nouc,"Test case passed for blank confirm password",14,
						"Courier New",true, true,IndexedColors.GREEN.getIndex(),
						IndexedColors.YELLOW.getIndex(),"CENTER");
				}
				else if(criteria.equalsIgnoreCase("email_alreadyused") && driver.findElement(
				By.xpath("//div[text()='This email is already used in an AidaForm account.']"))
					.isDisplayed())
				{
				Reporter.log("Test case passed for already used email");
				ef.setCellValue(sh,i,nouc,"Test case passed for already used email",14,
						"Courier New",true, true,IndexedColors.GREEN.getIndex(),
						IndexedColors.YELLOW.getIndex(),"CENTER");
				}
				else if(criteria.equalsIgnoreCase("pwd_wrongsize") && driver.findElement(By.xpath(
					"//div[text()='Enter a password longer than 6 characters']")).isDisplayed())
				{
				Reporter.log("Test case passed for wrong password size");
				ef.setCellValue(sh,i,nouc,"Test case passed for wrong password size",14,
						"Courier New",true, true,IndexedColors.GREEN.getIndex(),
						IndexedColors.YELLOW.getIndex(),"CENTER");
				}
				else if(criteria.equalsIgnoreCase("pwd_confirmpwd_different") && 
						driver.findElement(By.xpath("//div[text()=\"Passwords don't match\"]"))
						.isDisplayed())
				{
				Reporter.log("Test case passed for password mismatch");
				ef.setCellValue(sh,i,nouc,"Test case passed for password mismatch",14,
						"Courier New",true, true,IndexedColors.GREEN.getIndex(),
						IndexedColors.YELLOW.getIndex(),"CENTER");
				}
				else
				{
				Reporter.log(criteria+" Test case failed");
				ef.setCellValue(sh,i,nouc,criteria+" Test case failed",14,"Courier New",true, true,
						IndexedColors.RED.getIndex(),IndexedColors.YELLOW.getIndex(),"CENTER");
				//get screenshot and attach to TestNG report
				String ssfpath=ws.capturePageScreenshotAsFile(driver);
				Reporter.log(
				   "<a href=\""+ssfpath+"\"><img src=\""+ssfpath
				                       +"\" height=\"100\" width=\"100\"/></a>");
				//soft assert continues even if one test case fails
				sf.fail();	
				}
				//close site
				ws.closeSite(driver);
			}
			catch(Exception e)
			{
				Reporter.log(criteria+" Test case raised "+e.getMessage());
				ef.setCellValue(sh,i,nouc,
						criteria+" Test case raised "+e.getMessage().substring(0,50),
						14,	"Courier New",true, true,IndexedColors.RED.getIndex(),
						IndexedColors.YELLOW.getIndex(),"CENTER");
				//get screenshot and attach to TestNG report
				String ssfpath=ws.capturePageScreenshotAsFile(driver);
				Reporter.log(
				   "<a href=\""+ssfpath+"\"><img src=\""+ssfpath
				                       +"\" height=\"100\" width=\"100\"/></a>");
				//close site before hard stopping current test execution
				ws.closeSite(driver);
				Assert.fail();	//hard assert to stop current test execution
			}
		}
	}
	
	@AfterClass
	public void tearDown() throws Exception
	{
		//save and close excel file
		ef.saveAndCloseExcel(book,exceFilePath);
		sf.assertAll(); //Mandatory at end of code to see all test cases passed or failed
	}
}




