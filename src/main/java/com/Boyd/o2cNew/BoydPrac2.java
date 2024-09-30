package com.Boyd.o2cNew;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BoydPrac2 {
	
	
		static WebDriverWait wait;
		static WebDriver driver;
		public static int timeout = 60;
		
		@BeforeMethod()
			
		public void Logging() throws InterruptedException, IOException{
			
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			System.setProperty("webdriver.chrome.driver", "Users/iswarya.gumparthi_fo/Downloads/chromedriver_win32.exe");
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			driver.get("https://elme-dev1.fa.us8.oraclecloud.com/");
			driver.findElement(By.id("userid")).sendKeys("forsys.user");
			driver.findElement(By.id("password")).sendKeys("forsys2023");
			driver.findElement(By.id("btnActive")).click();
			WebElement homebutton = driver.findElement(By.id("pt1:_UIShome"));
			waitUntilElementClickable("homebutton", homebutton, driver, timeout);
			homebutton.click();
			Thread.sleep(10000);
			WebElement ordermanagement = driver.findElement(By.xpath("//div[@id='groupNode_order_management'][1]"));
			waitUntilElementClickable("ordermanagement", ordermanagement, driver, timeout);
			ordermanagement.click();
			WebElement ordermanagement1 = driver.findElement(By.xpath("//div[@id='itemNode_order_management_order_management']"));
			waitUntilElementClickable("ordermanagement1", ordermanagement1, driver, timeout);
			ordermanagement1.click();

		}
			
			@Test()
			
		public void OrderCreation() throws Exception
		{
			
		   File f=new File(System.getProperty("user.dir")+"\\ExcelData\\OrderCreation.xlsx");
				
			FileInputStream fis = new FileInputStream(f);

			XSSFWorkbook w = new XSSFWorkbook(fis);
			
			XSSFSheet sheet1 = w.getSheet("OrderCreation");
		
		
			sheet1.getRow(0).createCell(84).setCellValue("OrderNumber");
			sheet1.getRow(0).createCell(85).setCellValue("Result");
			sheet1.getRow(0).createCell(86).setCellValue("Comments");
			int totalrows= sheet1.getPhysicalNumberOfRows();
			System.out.println("Total no. of rows are:" +totalrows);

			
			for(int i=1;i<=totalrows;i++)
			{
				
				String Business_Unit = sheet1.getRow(i).getCell(0).getStringCellValue();
				String Customer = sheet1.getRow(i).getCell(1).getStringCellValue();
				String PurchaseOrder = sheet1.getRow(i).getCell(2).getStringCellValue();
				String OrderType = sheet1.getRow(i).getCell(3).getStringCellValue();
				String Contact = sheet1.getRow(i).getCell(4).getStringCellValue();
				String Contact_Method = sheet1.getRow(i).getCell(5).getStringCellValue();
				String Ship_to_Address = sheet1.getRow(i).getCell(6).getStringCellValue();
				String Bill_to_Customer = sheet1.getRow(i).getCell(7).getStringCellValue();
				String Bill_to_Address = sheet1.getRow(i).getCell(8).getStringCellValue();
				String PrimarySalesperson=sheet1.getRow(i).getCell(83).getStringCellValue();
				
				try
				{
				driver.findElement(By.xpath("//span[text()='Create Order']")).click();
			
				Thread.sleep(10000);
				WebElement bu = driver.findElement(By.xpath("//select[contains(@id, 'soc3::content')]"));
				bu.click();
				Select drpd = new Select(bu);
				drpd.selectByVisibleText(Business_Unit);
				
				WebElement customer=driver.findElement(By.id(
						"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:partyNameId::lovIconId"));
				waitUntilElementClickable("customer", customer, driver, timeout);
				customer.click();	
                Thread.sleep(10000);
				WebElement customer2=driver.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:partyNameId::dropdownPopup::popupsearch"));
				waitUntilElementClickable("customer2", customer2, driver, timeout);
				customer2.click();
				WebElement customer3=driver.findElement(By.xpath("//input[@id='pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:qryId1:value00::content'][1]"));
				waitUntilElementClickable("customer3", customer3, driver, timeout);
				customer3.sendKeys(Customer);
				
				WebElement customer4=driver.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:qryId1::search"));
				waitUntilElementClickable("customer4", customer4, driver, timeout);
				customer4.click();
				driver.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:partyNameId::lovDialogId::ok")).click();
				Thread.sleep(10000);
				WebElement purchase=driver.findElement(By.xpath("//input[contains(@id,'it1::content')]"));
				purchase.sendKeys(PurchaseOrder);
				Thread.sleep(10000);
				driver.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:orderType1Id::lovIconId")).click();
				WebElement Ordertype=driver.findElement(By.xpath("//span[text()='Standard']"));
				Ordertype.click();
				driver.findElement(By.xpath("//a[contains(@id,'AP1:primarySalesPersonNameId::lovIconId')]")).click();
				driver.findElement(By.xpath("//a[contains(@id,'AP1:primarySalesPersonNameId::dropdownPopup::popupsearch')]")).click();
				WebElement Salesperson=driver.findElement(By.xpath("//input[contains(@id,'AP1:qryId5:value00::content')]"));
				Salesperson.sendKeys(PrimarySalesperson);
				driver.findElement(By.xpath("//button[contains(@id,'AP1:qryId5::search')]")).click();
				driver.findElement(By.xpath("//button[contains(@id,'AP1:primarySalesPersonNameId::lovDialogId::ok')]")).click();
//		        driver.findElement(By.xpath("//span[text()='Save']")).click();
				Thread.sleep(10000);
			    driver.findElement(By.xpath("//a[contains(@id,'AP1:save::popEl')]")).click();
				driver.findElement(By.xpath("//span[text()='S']")).click();
				Thread.sleep(5000);
				System.out.println("Order Created succuesfully");
				
//				driver.findElement(By.xpath("//button[@id='pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:cb3']")).click();
				
				WebElement confirmation= driver.findElement(By.xpath("//td[contains(@id,'AP1:saveAndCloseDlg::contentContainer')]"));
				
				String orderconfirm= confirmation.getText();
				
				int OrderNumber= getNumericValue(orderconfirm);
				
				System.out.println("Order Number is:" +OrderNumber);

				driver.findElement(By.xpath("//button[@id='pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:cb3']")).click();
	
				sheet1.getRow(i).createCell(84).setCellValue(OrderNumber);
				
			    sheet1.getRow(i).createCell(85).setCellValue("Pass");
			    
			   sheet1.getRow(i).createCell(86).setCellValue("Order Created");
		       Updatefile(f, w);
				}
				catch(Exception e)
				{
//					 sheet1.getRow(i) .createCell(84).setCellValue("null");
//					   sheet1.getRow(i).createCell(85).setCellValue("fail");
//					   sheet1.getRow(i).createCell(86).setCellValue("Order is not Created");
//					   System.out.println("The provided values are Invaid");
//				       Updatefile(fis, xssf);
//				}
				}
			}
			}
//				  
			
			
				
			
			
			
			
			
			
			




			
				
			

			public void Updatefile(File f,XSSFWorkbook w)
			{
			try
			{
			FileOutputStream fos = new FileOutputStream(f);
			w.write(fos);
			fos.flush();
			}
			catch(Exception e)
			{
			e.printStackTrace();
			}
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
			
			
			
			public static void waitUntilElementClickable(String locatorName, final WebElement elementToWaitFor,
					WebDriver browser, int timeout) {
//				

							}


			}

			
