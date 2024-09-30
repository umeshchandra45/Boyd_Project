package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

@SuppressWarnings("unused")
public class Boyd_Order {
	public WebDriver browser;
	public int Order;
	public String Status;
	
	
	@BeforeTest
	public void Login_Page() throws Exception
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
		browser.get("https://elme.fa.us8.oraclecloud.com/");
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("Boyd2@2!");
		browser.findElement(By.id("btnActive")).click();
		Thread.sleep(12000);
		browser.findElement(By.linkText("You have a new home page!")).click();
		Thread.sleep(8000);
		browser.findElement(By.xpath("//a[text()='Order Management']")).click();
		browser.findElement(By.xpath("//*[contains(@id,'itemNode_order_management_order_management_1')]")).click();
		
		
		
	}
	
	@Test
	public void Homepage() throws Exception
	{
		File f = new File(System.getProperty("user.dir")+"\\Excel\\Boydorder.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		sheet.getRow(0).createCell(2).setCellValue("Result");
		sheet.getRow(0).createCell(3).setCellValue("Comments");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total rows of Excel is :" +totalrows);
		
		
		for(int i=1;i<=totalrows;i++)
		{
			if(sheet.getRow(i) == null)
			{
			return;
			}
			
			Order = (int)sheet.getRow(i).getCell(0).getNumericCellValue();
			String ordernum = String.valueOf(Order).trim();
			Status = sheet.getRow(i).getCell(1).getStringCellValue().trim();
			
			try {
			Thread.sleep(8000);
			browser.findElement(By.xpath("//a[text()='Tasks']")).click();
			browser.findElement(By.xpath("//td[text()='Manage Orders']")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("//input[contains(@id,'AP1:qryId1:value10::content')]")).click();
			browser.findElement(By.xpath("//input[contains(@id,'AP1:qryId1:value10::content')]")).sendKeys(ordernum);
			Thread.sleep(2000);
			Select sc = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId1:value70::content')]")));
			sc.selectByVisibleText(Status);
			browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId1::search')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("//a[text()='"+ordernum+"']")).click();
			Thread.sleep(10000);
			browser.findElement(By.xpath("//a[text()='Actions']")).click();
			try
			{
				browser.findElement(By.xpath("//td[text()='Create Revision']")).click();
			}
			catch(Exception e)
			{
				browser.findElement(By.xpath("//td[text()='Edit']")).click();
			}
			Thread.sleep(12000);
			try {
			browser.findElement(By.xpath("//span[text()='Save']")).click();
			Thread.sleep(12000);
			try {
			browser.findElement(By.xpath("//span[text()='Submit']")).click();
			try {
			WebDriverWait wait = new WebDriverWait(browser,250);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'APRS1:cb9')]")));
			browser.findElement(By.xpath("//*[contains(@id,'APRS1:cb9')]")).click();
			}
			catch(Exception e)
			{
				browser.findElement(By.xpath("//*[contains(@id,'AP1:cb14')]")).click();
			}
			Thread.sleep(4000);
			for(int k=1;k<=4;k++)
			{
				browser.findElement(By.xpath("//*[text()='Refresh']")).click();
				Thread.sleep(6000);
			}
			browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]")).click();
			Thread.sleep(4000);
			browser.findElement(By.xpath("//*[contains(@id,'AP1:SPb')]")).click();
			sheet.getRow(i).createCell(2).setCellValue("Pass");
	        Updatefile(f, wb);
			}
			catch(Exception e)
			{
				browser.findElement(By.xpath("//*[contains(@id,'APRS1:d4::ok')]")).click();
				Thread.sleep(3000);
				browser.findElement(By.xpath("//*[contains(@id,'APRS1:cancel')]")).click();
				Thread.sleep(6000);
				browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]")).click();
				Thread.sleep(4000);
				browser.findElement(By.xpath("//*[contains(@id,'AP1:SPb')]")).click();
				sheet.getRow(i).createCell(2).setCellValue("Fail");
				sheet.getRow(i).createCell(3).setCellValue("The order was not priced because the product charge for the item does not contain a value. Include a value, and then reprice the order.");
		        Updatefile(f, wb);
			}
			
			}
			catch(Exception e)
			{
				browser.findElement(By.xpath("//*[contains(@id,'APRS1:d18::ok')]")).click();
				Thread.sleep(4000);
				browser.findElement(By.xpath("//span[text()='ancel']")).click();
				WebDriverWait wait = new WebDriverWait(browser,250);
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[contains(@id,'APRS1:cb4')])[2]")));
				browser.findElement(By.xpath("(//*[contains(@id,'APRS1:cb4')])[2]")).click();
				Thread.sleep(3000);
				browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]")).click();
				Thread.sleep(4000);
				browser.findElement(By.xpath("//*[contains(@id,'AP1:SPb')]")).click();
				sheet.getRow(i).createCell(2).setCellValue("Fail");
				sheet.getRow(i).createCell(3).setCellValue("TPI attributes are not available.");
		        Updatefile(f, wb);
			}
			
			}
			
			catch(Exception e)
			{
				browser.findElement(By.xpath("//*[contains(@id,'AP1:SPb')]")).click();
				sheet.getRow(i).createCell(2).setCellValue("Fail");
				sheet.getRow(i).createCell(3).setCellValue("No data for given order");
		        Updatefile(f, wb);
			}
		}
		
		try {
	    	wb.close();
	    } catch(Exception e) {
	    	
	    }
		
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
	
	@AfterTest
	public void Quit_Browser()
	{
//		browser.quit();
	}
	

}
