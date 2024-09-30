package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

public class Boyd_Schedulesalesorder {
	
	public static WebDriver browser;
	public static int Order_Number;
	public String SSD_Date;
	

	
	
	@BeforeTest()
	public void Login_Page() throws Exception
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
//		browser.get("https://egmn-dev3.login.us2.oraclecloud.com/");
//		WebElement username = browser.findElement(By.id("userid"));
//		highLightElement(browser,username);
//		username.sendKeys("Jiong.tang@harmonicinc.com");
//		WebElement password = browser.findElement(By.id("password"));
//		highLightElement(browser,password);
//		password.sendKeys("welcome12345");
//		WebElement action = browser.findElement(By.id("btnActive"));
//		highLightElement(browser,action);
//		action.click();

		browser.get("https://elme-dev2.fa.us8.oraclecloud.com/");
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("Janakiram.Nalla");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("welcome1");
		browser.findElement(By.id("btnActive")).click();
		
		Thread.sleep(10000);
		WebElement homepage = browser.findElement(By.xpath("//a[text()='You have a new home page!']"));
		highLightElement(browser, homepage);
		homepage.click();
		Thread.sleep(7000);
		WebElement orderm = browser.findElement(By.xpath("//*[text()='Order Management']"));
		highLightElement(browser, orderm);
		orderm.click();
		WebElement orderm1 = browser.findElement(By.id("itemNode_order_management_order_management_1"));
		highLightElement(browser, orderm1);
		orderm1.click();
	}
	@Test
	public void Home_Page() throws Exception
	{
		
		File f = new File(System.getProperty("user.dir")+"\\Excel\\Scheduleorder.xlsx");
		FileInputStream fis = new FileInputStream(f);
		 XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		sheet.getRow(0).createCell(2).setCellValue("Result");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total rows of Excel is :" +totalrows);
		if(sheet.getRow(1).getCell(2) == null)
		{
		
		for(int i=1;i<=totalrows;i++)
		{
			if(sheet.getRow(i) == null)
			{
			return;
			}
			
		Order_Number = (int)sheet.getRow(i).getCell(0).getNumericCellValue();
		String number =String.valueOf(Order_Number);
		SSD_Date = sheet.getRow(i).getCell(1).getStringCellValue();
		
		Thread.sleep(6000);
		
		browser.findElement(By.xpath("//a[text()='Advanced']")).click();
		Thread.sleep(6000);
		Select sc = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId1:operator1::content')]")));
		sc.selectByVisibleText("Equals");
		Thread.sleep(2000);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId1:value10::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId1:value10::content')]")).sendKeys(number);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId1::search')]")).click();
		Thread.sleep(4000);
		browser.findElement(By.xpath("//a[text()='"+number+"']")).click();
		
		Thread.sleep(12000);
		WebElement actions = browser.findElement(By.xpath("//*[text()='Actions']"));
		highLightElement(browser, actions);
		actions.click();
		WebElement fullfilment = browser.findElement(By.xpath("//*[text()='Switch to Fulfillment View']"));
		highLightElement(browser, fullfilment);
		fullfilment.click();
		Thread.sleep(5000);
		WebElement fullfillment = browser.findElement(By.xpath("(//a[text()='Fulfillment Lines'])[1]"));
		highLightElement(browser, fullfillment);
		fullfillment.click();
		Thread.sleep(4000);
		WebElement table = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:DooFu1:0:ATT1:_ATTp:ATTt1::db')]/table/tbody/tr/td[1]"));
		highLightElement(browser, table);
		table.click();
		Thread.sleep(6000);
//		WebElement schedule = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:DooFu1:0:ATT1:_ATTp:ctb1')]"));
//		highLightElement(browser, schedule);
//		schedule.click();
//		Thread.sleep(10000);
//		WebElement ok = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:DooFu1:0:ATT1:d8::ok')]"));
//		highLightElement(browser, ok);
//		ok.click();
//		Thread.sleep(4000);
//		WebElement table1 = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:DooFu1:0:ATT1:_ATTp:ATTt1::db')]/table/tbody/tr/td[1]"));
//		highLightElement(browser, table1);
//		table1.click();
//		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'ATT1:_ATTp:edit::icon')]")).click();
		WebDriverWait wait = new WebDriverWait(browser,250);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'overrideScheduleDate::content')]")));
		Select sc1 = new Select(browser.findElement(By.xpath("//*[contains(@id,'overrideScheduleDate::content')]")));
		sc1.selectByVisibleText("Yes");
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'ATT1:DooFu2:1:id1::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'ATT1:DooFu2:1:id1::content')]")).sendKeys(SSD_Date);
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[text()='ave and Close']")).click();
		WebDriverWait wait1 = new WebDriverWait(browser,250);
		wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'ATT1:d8::ok')]")));
		browser.findElement(By.xpath("//*[contains(@id,'ATT1:d8::ok')]")).click();
		Thread.sleep(8000);
		for(int j=1;j<=4;j++)
		{
			WebElement rf1 = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:cb4')]"));
			highLightElement(browser, rf1);
			rf1.click();
			Thread.sleep(4000);
		}
		Thread.sleep(6000);
		String status = "Awaiting Shipping";
//		WebElement statusvalue = browser.findElement(By.xpath("//*[contains(@id,'ATT1:_ATTp:ATTt1::db')]/table/tbody/tr/td[2]/div/table/tbody/tr/td[21]"));
		WebElement statusvalue = browser.findElement(By.xpath("//*[contains(@id,'ATT1:_ATTp:ATTt1::db')]/table/tbody/tr/td[8]/div/table/tbody/tr/td[17]"));
		JavascriptExecutor js = (JavascriptExecutor)browser;
		js.executeScript("arguments[0].scrollIntoView()", statusvalue);
		String str = statusvalue.getText();
		System.out.println("Status value is :" +str);
		if(status.equalsIgnoreCase(str))
		{
			WebElement done = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:SPb')]"));
			highLightElement(browser, done);
			done.click();
			WebElement done1 = browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]"));
			highLightElement(browser, done1);
			done1.click();
			Thread.sleep(6000);
			WebElement done2 = browser.findElement(By.xpath("//*[contains(@id,'AP1:SPb')]"));
			highLightElement(browser, done2);
			done2.click();
		}
		else
		{
			for(int k=1;k<=4;k++)
			{
				WebElement rf1 = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:cb4')]"));
				highLightElement(browser, rf1);
				rf1.click();
				Thread.sleep(4000);
				WebElement done = browser.findElement(By.xpath("//*[contains(@id,'OrderDAP:SPb')]"));
				highLightElement(browser, done);
				done.click();
				WebElement done1 = browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]"));
				highLightElement(browser, done1);
				done1.click();
				Thread.sleep(6000);
				WebElement done2 = browser.findElement(By.xpath("//*[contains(@id,'AP1:SPb')]"));
				highLightElement(browser, done2);
				done2.click();
			}
		}
        sheet.getRow(i).createCell(2).setCellValue("Pass");
        Updatefile(f, wb);
		}
	}
		else
		{
			System.out.println("File is already uploaded");
		}
		
        
        try {
	    	wb.close();
	    } catch(Exception e) {
	    	
	    }
	
	}
	public static void highLightElement(WebDriver browser,WebElement ele)
	{
		try {  
            JavascriptExecutor js = (JavascriptExecutor) browser;  
            js.executeScript("arguments[0].style.border='4px groove red'", ele);  
            Thread.sleep(1000);  
            js.executeScript("arguments[0].style.border=''", ele);  
       } catch (Exception e) {  
            System.out.println(e);  
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
	@AfterTest()
	public void Close_Browser()
	{
//		browser.quit();
	}

}
