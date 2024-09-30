package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class Boyd_O2C_Scheduling_Sales_Order {
	public WebDriver browser;
	public String Order_Number;
	public String Fulfillment_Number;
	public String Item_Number;
	public String SSD_Update;
	
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
	//	browser.get("https://elme-dev2.fa.us8.oraclecloud.com");
		browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
//		browser.get("https://elme-test.login.us8.oraclecloud.com/");
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("forsys2023");
		browser.findElement(By.id("btnActive")).click();
		Thread.sleep(5000);
		WebDriverWait wait1 = new WebDriverWait(browser, 500);
		wait1.until(ExpectedConditions.elementToBeClickable(By.id("pt1:_UIShome")));
		browser.findElement(By.id("pt1:_UIShome")).click();
		Thread.sleep(26000);
		browser.findElement(By.linkText("Order Management")).click();
		WebElement order1 = browser.findElement(By.id("itemNode_order_management_order_management_1"));
		WebDriverwaitelement(order1);
		order1.click();
	}
	@Test()
	public void Home_Page() throws Exception
	{
		
		File f = new File(System.getProperty("user.dir")+"\\Excel\\Boyd_O2C_Scheduling_SalesOrder.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Schedulingsalesorder");
		sheet.getRow(0).createCell(4).setCellValue("Result");
		sheet.getRow(0).createCell(5).setCellValue("Comments");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		
		if(sheet.getRow(1).getCell(4) == null)
		{
		
		for(int i=1;i<=totalrows;i++)
		{
			
			if(sheet.getRow(i) == null)
			{
				return;
			}
			
			Order_Number = sheet.getRow(i).getCell(0).getStringCellValue();
			Fulfillment_Number = sheet.getRow(i).getCell(1).getStringCellValue();
			Item_Number = sheet.getRow(i).getCell(2).getStringCellValue();
			SSD_Update = sheet.getRow(i).getCell(3).getStringCellValue();
			
			Thread.sleep(4000);
			WebElement task = browser.findElement(By.linkText("Tasks"));
			WebDriverwaitelement(task);
			task.click();
			WebElement fulfillment = browser.findElement(By.xpath("//td[text()='Manage Fulfillment Lines']"));
			WebDriverwaitelement(fulfillment);
			fulfillment.click();
			Thread.sleep(8000);
			WebElement el = browser.findElement(By.xpath("//*[contains(@id,'value20::content')]"));
			WebDriverwaitelement(el);
			el.click();
			Select sc = new Select(browser.findElement(By.xpath("//*[contains(@id,'operator2::content')]")));
			sc.selectByVisibleText("Equals");
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'value20::content')]")).click();
			Thread.sleep(2000);
			browser.findElement(By.xpath("//*[contains(@id,'value20::content')]")).sendKeys(Order_Number);
			Thread.sleep(2000);
			browser.findElement(By.xpath("//*[contains(@id,'value30::content')]")).click();
			browser.findElement(By.xpath("//*[contains(@id,'value30::content')]")).sendKeys(Fulfillment_Number);
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'value50::content')]")).click();
			browser.findElement(By.xpath("//*[contains(@id,'value50::content')]")).sendKeys(Item_Number);
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'q1::search')]")).click();
			try
			{
			WebElement table = browser.findElement(By.xpath("//*[contains(@id,'ATt1::db')]/table/tbody/tr/td[1]"));
			WebDriverwaitelement(table);
			table.click();
			Thread.sleep(3000);
			WebElement edit = browser.findElement(By.xpath("//*[contains(@id,'edit::icon')]"));
			JavascriptExecutor js = (JavascriptExecutor)browser;
			js.executeScript("arguments[0].click();", edit);
			Thread.sleep(8000);
			Select overide = new Select(browser.findElement(By.xpath("//*[contains(@id,'overrideScheduleDate::content')]")));
			overide.selectByVisibleText("Yes");
			Thread.sleep(6000);
			browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).click();
			 DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
			 Calendar cal = Calendar.getInstance();
			browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).sendKeys(dateFormat.format(cal.getTime())+" 09:15 PM");
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'FulSAP:AT1:cb4')]")).click();
			try
			{
			WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'FulSAP:AT1:d9::ok')]"));
			WebDriverwaitelement(okbutton);
			okbutton.click();
			Thread.sleep(3000);
			 WebElement refresh = browser.findElement(By.xpath("//button[text()='Refresh']"));
			 WebDriverwaitelement(refresh);
			 refresh.click();
			 Thread.sleep(4000);
			 browser.findElement(By.xpath("//button[text()='Refresh']")).click();
			 Thread.sleep(3000);
			 browser.findElement(By.xpath("//button[text()='Refresh']")).click();
			 Thread.sleep(6000);
			 browser.findElement(By.xpath("//*[contains(@id,'FulSAP:cb1')]")).click();
			 sheet.getRow(i).createCell(4).setCellValue("Pass");
			 Updatefile(f, wb);
			 
			}
			catch(Exception e)
			{
				browser.findElement(By.id("d1::msgDlg::cancel")).click();
				WebElement cancel = browser.findElement(By.xpath("//*[contains(@id,'d3::cancel')]"));
				WebDriverwaitelement(cancel);
				cancel.click();
				WebElement done = browser.findElement(By.xpath("//*[contains(@id,'FulSAP:cb1')]"));
				WebDriverwaitelement(done);
				done.click();
				sheet.getRow(i).createCell(4).setCellValue("Fail");
				sheet.getRow(i).createCell(5).setCellValue("You cannot set the scheduled ship date to a date prior to today.");
				Updatefile(f, wb);
			}
		}
			catch(Exception e)
			{
				browser.findElement(By.xpath("//*[contains(@id,'FulSAP:cb1')]")).click();
				sheet.getRow(i).createCell(4).setCellValue("Fail");
				sheet.getRow(i).createCell(5).setCellValue("Order has no data or edit button is in disablemode");
				Updatefile(f, wb);
				
			}
		}
	}
		else
		{
			System.out.println("File is already processed");
		}
		try
		{
			wb.close();
		}
		catch(Exception e)
		{
			
		}
		
	}
	
	
	public void WebDriverwaitelement(WebElement element)
	{
		WebDriverWait wait = new WebDriverWait(browser,350);
		wait.until(ExpectedConditions.visibilityOf(element));
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
	public void Close_browser()
	{
		browser.quit();
	}

}
