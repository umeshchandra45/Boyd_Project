package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

@SuppressWarnings("unused")
public class BOYD_O2C_Account_Creation {
	public WebDriver browser;
	public String Subledger_Application;
	public String Ledger;
	public String Process_Category;
	public String End_date;
	public String Accounting_Mode;
	public String Process_Events;
	public String Report_Style;
	public String Transfer_to_GL;
	public String Post_in_GL;
	public static WebDriverWait wait;
	public static int timeout = 60;
	
	
	
	
	@BeforeTest()
	public void Login_Page() throws Exception
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
	//browser.get("https://elme-dev2.fa.us8.oraclecloud.com");
		browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
//		browser.get("https://elme-test.login.us8.oraclecloud.com/");
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("forsys2023");
	//	browser.findElement(By.id("password")).sendKeys("forsys4@4!");
		browser.findElement(By.id("btnActive")).click();
		WebElement homepage = browser.findElement(By.xpath("//a[text()='You have a new home page!']"));
		waitUntilElementClickable("homepage", homepage, browser, timeout);
		WebElement tools = browser.findElement(By.linkText("Tools"));
		waitUntilElementClickable("tools", tools, browser, timeout);
		WebElement scheduledprocess = browser.findElement(By.linkText("Scheduled Processes"));
		waitUntilElementClickable("scheduledprocess", scheduledprocess, browser, timeout);
		
	}
	@Test()
	public void Home_Page() throws Exception
	{
		File f = new File(System.getProperty("user.dir")+"\\Excel\\BOYD_O2C_Create_Accounting.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Account_Creation");
		sheet.getRow(0).createCell(9).setCellValue("Process_ID");
		sheet.getRow(0).createCell(10).setCellValue("Result");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		
		if(sheet.getRow(1).getCell(10) == null)
		{
		for(int i=1;i<=totalrows;i++)
		{
			
			if(sheet.getRow(i) == null)
			{
				return;
			}
			
			
			Subledger_Application = sheet.getRow(i).getCell(0).getStringCellValue();
			Ledger = sheet.getRow(i).getCell(1).getStringCellValue();
			Process_Category = sheet.getRow(i).getCell(2).getStringCellValue();
			End_date = sheet.getRow(i).getCell(3).getStringCellValue();
			Accounting_Mode = sheet.getRow(i).getCell(4).getStringCellValue();
			Process_Events = sheet.getRow(i).getCell(5).getStringCellValue();
			Report_Style = sheet.getRow(i).getCell(6).getStringCellValue();
			Transfer_to_GL = sheet.getRow(i).getCell(7).getStringCellValue();
			Post_in_GL = sheet.getRow(i).getCell(8).getStringCellValue();
			
			WebElement task = browser.findElement(By.linkText("Schedule New Process"));
			waitUntilElementClickable("task", task, browser, timeout);
			WebElement searchdropdown = browser.findElement(By.xpath("//*[contains(@id,'selectOneChoice2::lovIconId')]"));
			waitUntilElementClickable("searchdropdown", searchdropdown, browser, timeout);
			WebElement searchicon = browser.findElement(By.linkText("Search..."));
			waitUntilElementClickable("searchicon", searchicon, browser, timeout);
			WebElement name = browser.findElement(By.xpath("//*[contains(@id,'_afrLovInternalQueryId:value00::content')]"));
			waitUntilElementClickable("name", name, browser, timeout);
			name.sendKeys("Create Accounting");
			WebElement searchbutton = browser.findElement(By.xpath("//*[contains(@id,'_afrLovInternalQueryId::search')]"));
			waitUntilElementClickable("searchbutton", searchbutton, browser, timeout);
			WebElement table1 = browser.findElement(By.xpath("//*[contains(@id,'selectOneChoice2_afrLovInternalTableId::db')]/table/tbody/tr[1]/td[1]"));
			waitUntilElementClickable("table1", table1, browser, timeout);
			WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'lovDialogId::ok')]"));
			waitUntilElementClickable("okbutton", okbutton, browser, timeout);
			Thread.sleep(4000);
		    WebElement button = browser.findElement(By.xpath("//*[contains(@id,'snpokbtnid')]"));
		    waitUntilElementClickable("button", button, browser, timeout);
			Thread.sleep(8000);
			Select sc = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_SubledgerApplicationAttr::content')]")));
			sc.selectByVisibleText(Subledger_Application);
			WebElement buttonvalue = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_LedgerAttr::lovIconId')]"));
			waitUntilElementClickable("buttonvalue", buttonvalue, browser, timeout);
			WebElement search = browser.findElement(By.linkText("Search..."));
			waitUntilElementClickable("search", search, browser, timeout);
			WebElement ledger = browser.findElement(By.xpath("//*[contains(@id,'_afrLovInternalQueryId:value00::content')]"));
			waitUntilElementClickable("ledger", ledger, browser, timeout);
			WaituntilElementwritable("ledger", ledger, browser, Ledger);
			WebElement searchvalue = browser.findElement(By.xpath("//button[text()='Search']"));
			waitUntilElementClickable("searchvalue", searchvalue, browser, timeout);
			WebElement table = browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_LedgerAttr_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
			waitUntilElementClickable("table", table, browser, timeout);
			WebElement okvalue = browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_LedgerAttr::lovDialogId::ok')]"));
			waitUntilElementClickable("okvalue", okvalue, browser, timeout);
			Thread.sleep(6000);
			Select process = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_ATTRIBUTE5_ATTRIBUTE5::content')]")));
			process.selectByVisibleText(Process_Category);
			Thread.sleep(4000);
			WebElement enddate = browser.findElement(By.xpath("//input[contains(@id,'basicReqBody:paramDynForm_ATTRIBUTE6::content')]"));
			waitUntilElementClickable("enddate", enddate, browser, timeout);
			browser.findElement(By.xpath("//input[contains(@id,'basicReqBody:paramDynForm_ATTRIBUTE6::content')]")).clear();
			WaituntilElementwritable("enddate", enddate, browser, End_date);
			browser.findElement(By.xpath("//input[contains(@id,'basicReqBody:paramDynForm_ATTRIBUTE6::content')]")).sendKeys(Keys.ENTER);
			Thread.sleep(2000);
			Select mode = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_ATTRIBUTE8_ATTRIBUTE8::content')]")));
			mode.selectByVisibleText(Accounting_Mode);
			Select event = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_ATTRIBUTE9_ATTRIBUTE9::content')]")));
			event.selectByVisibleText(Process_Events);
			Select style = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_ATTRIBUTE10_ATTRIBUTE10::content')]")));
			style.selectByVisibleText(Report_Style);
			Select gl = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_ATTRIBUTE11_ATTRIBUTE11::content')]")));
			gl.selectByVisibleText(Transfer_to_GL);
			Select pgl = new Select(browser.findElement(By.xpath("//*[contains(@id,'paramDynForm_ATTRIBUTE12_ATTRIBUTE12::content')]")));
			pgl.selectByVisibleText(Post_in_GL);
			Thread.sleep(6000);
			WebElement ele = browser.findElement(By.xpath("//*[contains(@id,'notifyOption::content')]"));
			JavascriptExecutor js = (JavascriptExecutor)browser;
			js.executeScript("arguments[0].click()",ele );
			WebElement submitbutton = browser.findElement(By.xpath("//*[contains(@id,'requestBtns:submitButton')]"));
			waitUntilElementClickable("submitbutton", submitbutton, browser, timeout);
			Thread.sleep(5000);
			WebElement conform = browser.findElement(By.xpath("//span[contains(@id,'requestBtns:confirmationPopup:pt_ol1')]"));
			String order = conform.getText();
			System.out.println("The order text is :"+order);
			Thread.sleep(3000);
			int value = getNumericValue(order);
			String str = String.valueOf(value);
			System.out.println("The order value is :" +value);
			sheet.getRow(i).createCell(9).setCellValue(value);
			Updatefile(f, wb);
			WebElement ok = browser.findElement(By.xpath("//*[contains(@id,'requestBtns:confirmationPopup:confirmSubmitDialog::ok')]"));
			waitUntilElementClickable("ok", ok, browser, timeout);
			WebElement done = browser.findElement(By.xpath("//*[contains(@id,'srRssdfl::_afrDscl')]"));
			waitUntilElementClickable("done", done, browser, timeout);
			WebElement ordervalue = browser.findElement(By.xpath("//*[contains(@id,'srRssdfl:value10::content')]"));
			waitUntilElementClickable("ordervalue", ordervalue, browser, timeout);
			browser.findElement(By.xpath("//*[contains(@id,'srRssdfl:value10::content')]")).sendKeys(str.trim());
			WebElement search1 = browser.findElement(By.xpath("//*[contains(@id,'srRssdfl::search')]"));
			waitUntilElementClickable("search1", search1, browser, timeout);
			for(int k=1;k<=5;k++)
			{
				browser.findElement(By.xpath("//*[contains(@id,'panel:processRefreshId::icon')]")).click();
				Thread.sleep(6000);
			}
			Thread.sleep(6000);
			WebElement homebutton = browser.findElement(By.id("pt1:_UIShome::icon"));
			waitUntilElementClickable("homebutton", homebutton, browser, timeout);
			WebElement toolbutton = browser.findElement(By.linkText("Tools"));
			waitUntilElementClickable("toolbutton", toolbutton, browser, timeout);
			WebElement order1 = browser.findElement(By.linkText("Scheduled Processes"));
			waitUntilElementClickable("order1", order1, browser, timeout);
			sheet.getRow(i).createCell(10).setCellValue("Pass");
			Updatefile(f, wb);
			
		}
		
		
	}	
		else
		{
		   System.out.println("File is already Processed");
		}
		
		try
		{
			wb.close();
		}
		catch(Exception e)
		{
			
		}
	}
	
	
	public static void waitUntilElementClickable(String locatorName, final WebElement elementToWaitFor,
			WebDriver browser, int timeout) {
//		System.out.println("<<<<<< "+locatorName+">>>>>>>>");
		wait = new WebDriverWait(browser, timeout);
		wait.until(new Function<WebDriver, Boolean>() {
			int j;

			public Boolean apply(WebDriver browser) {
				j++;
				if (elementToWaitFor.isEnabled()) {
					try {
						elementToWaitFor.click();

					} catch (Exception e) {
						return false;

					}

				}
				return true;

			}
		});

	}
	
	public static void WaituntilElementwritable(String locatorName, final WebElement elementToWaitFor,
			WebDriver browser, String value) {
//		System.out.println("<<<<<< "+locatorName+" >>>>>>>>");

		wait = new WebDriverWait(browser, timeout);
		wait.until(new Function<WebDriver, Boolean>() {
			int j;

			public Boolean apply(WebDriver browser) {
				j++;
				if (elementToWaitFor.isEnabled()) {
					try {
						elementToWaitFor.sendKeys(value);

					} catch (Exception e) {
						return false;

					}

				}
				return true;

			}
		});

	}
	
	public static Integer getNumericValue(String str) {
		String str1[] = str.split("\\s");
		for (String s : str1) {
		boolean isNumeric = s.trim().chars().allMatch(Character::isDigit);
		if (isNumeric) {
		return Integer.parseInt(s);
		}
		}
		return 0;
		}
	
	public void Updatefile(File f,XSSFWorkbook wb)
	{
	try
	{
	FileOutputStream fos = new FileOutputStream(f);
	wb.write(fos);
	fos.flush();
	}
	 catch(Exception e)
	{
	e.printStackTrace();
	}
	}
	
	
	@AfterTest()
	public void Close_Browser()
	{
//		browser.quit();
	}

}
