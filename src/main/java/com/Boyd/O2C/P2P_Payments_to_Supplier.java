package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Set;
import java.util.concurrent.TimeUnit;

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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class P2P_Payments_to_Supplier {
	public WebDriver driver;
	public String BUName;
	public String supplier;
	public String Invoice_number;
	public String Priority_Override;
	public String Comments;
	public String Bank_Account;
	public String Business_Unit;
	public String Legal_Entity;
	public String Payment_Process_profile;
	public String Payment_Document;
	public String Transmit_Now;
	
	@BeforeTest()
	public void Login_Page() throws Exception
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		driver = new ChromeDriver(options);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
		driver.get("https://elme-dev1.fa.us8.oraclecloud.com");
		driver.findElement(By.xpath("//input[contains(@id, 'userid')]")).sendKeys("forsys.user");
		Thread.sleep(2000);
		driver.findElement(By.xpath("//input[contains(@id, 'password')]")).sendKeys("forsys2023");
		driver.findElement(By.id("btnActive")).click();
		Thread.sleep(22000);
		driver.findElement(By.id("pt1:_UIShome::icon")).click();
		Thread.sleep(15000);
//		driver.findElement(By.xpath("//*[contains(@id,'clusters-right-nav')]")).click();
//		Thread.sleep(2000);
//		driver.findElement(By.xpath("//*[contains(@id,'clusters-right-nav')]")).click();
//		Thread.sleep(2000);
//		driver.findElement(By.xpath("//*[contains(@id,'clusters-right-nav')]")).click();
//		Thread.sleep(2000);
		driver.findElement(By.linkText("Payables")).click();
		WebElement payments = driver.findElement(By.linkText("Payments"));
		WebDriverwaitelement(payments);
		payments.click();
	}
	
	@SuppressWarnings("unused")
	@Test()
	public void Home_Page() throws Exception
	{
		
		File f = new File(System.getProperty("user.dir")+"\\Excel\\P2P_Cycle.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Payments_to_Supplier");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		for(int i=1;i<=totalrows;i++)
		{
			if(sheet.getRow(i) == null)
			{
				return;
			}
			
			BUName = sheet.getRow(i).getCell(0).getStringCellValue();
			supplier = sheet.getRow(i).getCell(1).getStringCellValue();
			Invoice_number = sheet.getRow(i).getCell(2).getStringCellValue();
			Bank_Account = sheet.getRow(i).getCell(3).getStringCellValue();	
			Payment_Process_profile = sheet.getRow(i).getCell(4).getStringCellValue();
			Payment_Document = sheet.getRow(i).getCell(5).getStringCellValue();
			
			
			
			Thread.sleep(6000);
			WebElement task = driver.findElement(By.xpath("//*[contains(@id,'_FOTsdi__PaymentLanding_itemNode__FndTasksList::icon')]"));
			WebDriverwaitelement(task);
			task.click();
			WebElement paymentprocess = driver.findElement(By.linkText("Create Payment"));
			WebDriverwaitelement(paymentprocess);
			paymentprocess.click();
			WebElement bU = driver.findElement(By.xpath("//*[contains(@id,'OrgUiId::content')]"));
			WebDriverwaitelement(bU);
			bU.clear();
			bU.click();
			bU.sendKeys(BUName);
			bU.sendKeys(Keys.ENTER);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//*[contains(@id,'payeeNameId::lovIconId')]")).click();
			Thread.sleep(2000);
			WebElement supplierfield = driver.findElement(By.xpath("//*[contains(@id,'LovInternalQueryId:value00::content')]"));
			WebDriverwaitelement(supplierfield);
			supplierfield.click();
			supplierfield.clear();
			supplierfield.sendKeys(supplier);
			Thread.sleep(2000);
			driver.findElement(By.xpath("//button[text()='Search']")).click();
			WebElement tablerow = driver.findElement(By.xpath("//*[contains(@id,'payeeNameId_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
			WebDriverwaitelement(tablerow);
			tablerow.click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//*[contains(@id,'payeeNameId::lovDialogId::ok')]")).click();
			Thread.sleep(4000);
//			try {
//			driver.findElement(By.xpath("//*[contains(@id,'cb10')]")).click();
//			Thread.sleep(2000);
//			}
//			catch(Exception e) {
//				System.out.println("Ok button not found");
//			}
			
	    	WebElement bankaccount=	driver.findElement(By.xpath("//*[contains(@id,'bankAccountNameId::content')]"));
	    	WebDriverwaitelement(bankaccount);
	    	bankaccount.click();
	    	bankaccount.clear();
	    	bankaccount.sendKeys(Bank_Account);
	    	bankaccount.sendKeys(Keys.ENTER);
	    	Thread.sleep(6000);
	    	WebElement PaymentMethod=	driver.findElement(By.xpath("//input[contains(@id,'paymentMethodNameUiId::content')]"));
	    	WebDriverwaitelement(PaymentMethod);
	    	PaymentMethod.click();
	    	PaymentMethod.clear();
	    	PaymentMethod.sendKeys("Wire");
	    	Thread.sleep(5000);
	    	PaymentMethod.sendKeys(Keys.ENTER);
	    	Thread.sleep(4000);
			WebElement el =  driver.findElement(By.xpath("//*[contains(@id,'paymentProfileUICompId::content')]"));
			WebDriverwaitelement(el);
            el.click();
            el.clear();
            el.sendKeys(Payment_Process_profile);
            el.sendKeys(Keys.ENTER);
            Thread.sleep(5000);
            driver.findElement(By.xpath("//*[contains(@id,'PaymentDocumentIdUi::lovIconId')]")).click();
            Thread.sleep(1000);
            WebElement searchicon = driver.findElement(By.xpath("//a[contains(@id,'PaymentDocumentIdUi::dropdownPopup::popupsearch')]"));
			WebDriverwaitelement(searchicon);
			searchicon.click();
			WebElement documentfield = driver.findElement(By.xpath("//*[contains(@id,'PaymentDocumentIdUi::_afrLovInternalQueryId:value00::content')]"));
			WebDriverwaitelement(documentfield);
			documentfield.click();
			documentfield.clear();
			documentfield.sendKeys(Payment_Document);
			Thread.sleep(2000);
			driver.findElement(By.xpath("//button[text()='Search']")).click();
			WebElement tablerow1 = driver.findElement(By.xpath("//*[contains(@id,'PaymentDocumentIdUi_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
			WebDriverwaitelement(tablerow1);
			tablerow1.click();
			Thread.sleep(3000);
			driver.findElement(By.xpath("//*[contains(@id,'PaymentDocumentIdUi::lovDialogId::ok')]")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//img[contains(@id,'ATp:commandToolbarButton1::icon')]")).click();
			Thread.sleep(3000);
			WebElement invInput= driver.findElement(By.xpath("//input[contains(@id,'coVOId:value00::content')]"));
			WebDriverwaitelement(invInput);
			invInput.clear();
			invInput.sendKeys(Invoice_number);
			Thread.sleep(2000);
			driver.findElement(By.xpath("//button[text()='Search']")).click();
			Thread.sleep(2000);
			WebElement invtablerow1 = driver.findElement(By.xpath("//*[contains(@id,'ResultId:_ATp:t1::db')]/table/tbody/tr[1]/td[1]"));
			WebDriverwaitelement(invtablerow1);
			invtablerow1.click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//button[contains(text(),'App')]")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//button[contains(@id,'dialog1::ok')]")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//button[text()='ave and Close']")).click();
			Thread.sleep(10000);
			WebElement Paynum = driver.findElement(By.xpath("//*[@id='d1::msgDlg::_cnt']/div/table/tbody/tr/td/table/tbody/tr/td[2]/div"));
			String str =Paynum.getText().trim();
			Thread.sleep(3000);
			String[] str3=str.split(" ");
			Thread.sleep(3000);
			String str2 = str3[1];
			System.out.println("payment num=="+str2);
			sheet.getRow(i).createCell(6).setCellValue(str2);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//button[@id='d1::msgDlg::cancel']")).click();
			Thread.sleep(3000);
		    sheet.getRow(i).createCell(7).setCellValue("Pass");
			Updatefile(f, wb);
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
		WebDriverWait wait = new WebDriverWait(driver,350);
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
	public void Close_driver()
	{
//		driver.quit();
	}
	
	

}
