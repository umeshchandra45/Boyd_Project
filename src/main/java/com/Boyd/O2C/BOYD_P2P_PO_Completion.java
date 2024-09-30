package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BOYD_P2P_PO_Completion {
	
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	    WebDriver driver;
	    String result;
		File f;
		 int i;
		 XSSFSheet sheet;
	   
	    @SuppressWarnings("unused")
		@Test
	public void purchaseOrder() throws InterruptedException, IOException
	{
	try {
		WebDriverManager.chromedriver().setup();
	 driver = new ChromeDriver();
	 driver.manage().deleteAllCookies();
	 driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
	 driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
//	 driver.get("https://elme-dev1.fa.us8.oraclecloud.com");
	 driver.manage().window().maximize();

	 
//	 String remoteExecution = System.getProperty("remoteExecution");
//		System.out.println("remoteFlag  is :" +remoteExecution);
//		Boolean exectionFlag = Boolean.parseBoolean(remoteExecution);
		File f = null;
////		if(exectionFlag) {
//			System.out.println("remote execution");
//			f = new File(System.getProperty("user.dir") + "\\ExternalFiles\\BOYD_P2P_PurchaseOrder_Completion.xlsx");
//		}			 
////		else {
			 f = new File(System.getProperty("user.dir") + "\\Excel\\BOYD_P2P_PurchaseOrder_Completion.xlsx");
//			 System.out.println("local execution");
//		}
	 
	 fis = new FileInputStream(f);
	 wb = new XSSFWorkbook(fis);
	 XSSFSheet sheet = wb.getSheet("PO_Completion");
	//browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
//			String url = System.getProperty("url");
//			System.out.println("url is :" +url);
//			String username = System.getProperty("userName");
//			System.out.println("username  is :" +username);
//			String password = System.getProperty("password");
//			System.out.println("password  is :" +password);
			
//			driver.get("https://elme-dev2.fa.us8.oraclecloud.com");
//	 driver.get("https://elme-test.login.us8.oraclecloud.com/");
	 driver.get("https://elme-dev1.fa.us8.oraclecloud.com");
			driver.findElement(By.id("userid")).click();	
			//browser.findElement(By.id("userid")).sendKeys("forsys.user");
			
			
			driver.findElement(By.id("userid")).sendKeys("forsys.user");
			driver.findElement(By.id("password")).click();
			//browser.findElement(By.id("password")).sendKeys("welcome123");	
			//driver.findElement(By.id("password")).sendKeys("forsys2023");

			driver.findElement(By.id("password")).sendKeys("forsys2023");
	 
	 JavascriptExecutor js = (JavascriptExecutor) driver;
	 driver.findElement(By.id("btnActive")).click();
	 Thread.sleep(5000);
	  WebDriverWait wait1 = new WebDriverWait(driver, 500);
	  wait1.until(ExpectedConditions.elementToBeClickable(By.id("pt1:_UIShome")));
	  driver.findElement(By.id("pt1:_UIShome")).click();
	 Thread.sleep(6000);
	 js.executeScript("window.scrollBy(0, 2000)", "");
	 Thread.sleep(10000);
	 driver.findElement(By.linkText("Procurement")).click();
	 driver.findElement(By.linkText("Purchase Orders")).click();
	 Thread.sleep(5000);
	 int rowNum=sheet.getPhysicalNumberOfRows();
	 System.out.println("rowNum="+rowNum);
	 System.out.println("colNum="+sheet.getRow(1).getLastCellNum());
	 Row row = sheet.getRow(1);
	  Cell c = row.getCell(21);
	  System.out.println("result=="+c);
	 if(c==null||c.getStringCellValue().contentEquals(""))
		 {
	 for(int i=1; i<=rowNum; i++)
	 {
		 if(sheet.getRow(i) == null || isRowEmpty(sheet.getRow(i))) {
			 
			 driver.findElement(By.xpath("//span[text()='Save']")).click();
				Thread.sleep(15000);
				driver.findElement(By.xpath("//div[contains(@id, 'AP1:SPsb2')]")).click();
				Thread.sleep(7000);
				WebDriverWait wait11 = new WebDriverWait(driver, 350);
				wait11.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"d1::msgDlg::_cnt\"]/div/table/tbody/tr/td/table/tbody/tr/td[2]/div")));
				WebElement POnum = driver.findElement(By.xpath("//*[@id=\"d1::msgDlg::_cnt\"]/div/table/tbody/tr/td/table/tbody/tr/td[2]/div"));
				
				String str =POnum.getText().trim();
				Thread.sleep(3000);
				String[] str3=str.split(" ");
				Thread.sleep(3000);
				String str2 = str3[4];
				System.out.println("PO num=="+str2);
				sheet.getRow(i-1).createCell(19).setCellValue(str2);
				fos = new FileOutputStream(f);
			    wb.write(fos);
			    driver.findElement(By.xpath("//button[contains(@id, 'd1::msgDlg::cancel')]")).click();
			    Thread.sleep(5000);
				driver.findElement(By.xpath("//img[contains(@id, 'FndTasksList::icon')]")).click();
				Thread.sleep(4000);
				driver.findElement(By.linkText("Manage Orders")).click();
				Thread.sleep(5000);
				WebElement bu = driver.findElement(By.xpath("//select[contains(@id, 'value10::content')]"));
				 Select bussUnt = new Select(bu);
				 bussUnt.selectByIndex(0);
				 Thread.sleep(4000);
				 driver.findElement(By.xpath("//input[contains(@id, 'value40::content')]")).click();
				 driver.findElement(By.xpath("//input[contains(@id, 'value40::content')]")).clear();
		        driver.findElement(By.xpath("//input[contains(@id, 'value40::content')]")).sendKeys(str2);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//button[contains(@id, 'search')]")).click();
		        Thread.sleep(15000);
				String status = driver.findElement(By.xpath("//span[contains(@id, 'AP1:r1:0:AT1:_ATp:table1:0:ot49')]")).getText();
		       System.out.println("status=="+status);
		       if(status.equalsIgnoreCase("Pending Approval"))
		       {
////		    	   Thread.sleep(6000);
////		        driver.findElement(By.linkText("Pending Approval")).click();
////		        Thread.sleep(5000);
////		        js.executeScript("window.scrollBy(0, 2000)", "");
//		        driver.findElement(By.xpath("(//img[contains(@id, 'snapshot::icon')])[2]")).click();
//		        Thread.sleep(6000);
//				 Set<String> windows =  driver.getWindowHandles();
//				  System.out.println("no.of windows=" +windows.size());
//				  String [] array = windows.toArray(new String[windows.size()]);
//				  
//		           String window1 = array[0];
//		           String window2 = array[1];
//		           driver.switchTo().window(window2);
//		           Thread.sleep(5000);
//		           driver.findElement(By.xpath("//button[text()='Approve']")).click();
//		           Thread.sleep(5000);
//		           driver.findElement(By.xpath("//*[contains(@id, 'apprej::content')]")).sendKeys("Approved");
//		           driver.findElement(By.xpath("//*[contains(@id, 'apprej_submit')]")).click();
//		           Thread.sleep(4000);
//		           driver.switchTo().window(window2);
//		           Thread.sleep(4000);
//		           driver.close();
//		           driver.switchTo().window(window1);
//		           Thread.sleep(5000);
//		           driver.findElement(By.xpath("//div[contains(@id, 'AP1:SPb')]")).click();
//			   }
//		           Thread.sleep(5000);
//		           driver.findElement(By.xpath("//div[contains(@id, 'AP1:SPb')]")).click();
//		           Thread.sleep(5000);
		           CellStyle style = wb.createCellStyle();
			  		 Font font = wb.createFont();
			  		 XSSFCell cell2 = sheet.getRow(i-1).createCell(20);
			  		 cell2.setCellValue("Pass");
			  		 font.setColor(IndexedColors.GREEN.getIndex());
			  		 font.setBold(true);
			  		 style = wb.createCellStyle();
			  		 style.setFont(font);
			  		 cell2.setCellStyle(style);
			  		 fos = new FileOutputStream(f);
			  		 wb.write(fos);
			  		continue;
		 
		 }
		 }
	  String procumentBU = sheet.getRow(i).getCell(0).getStringCellValue().trim();
	  String requisionBU = sheet.getRow(i).getCell(1).getStringCellValue().trim();
	  String purStyle = sheet.getRow(i).getCell(2).getStringCellValue().trim();
	  String supplier = sheet.getRow(i).getCell(3).getStringCellValue().trim();
	  String supplier_site = sheet.getRow(i).getCell(4).getStringCellValue().trim();
	  String shipLoc = sheet.getRow(i).getCell(5).getStringCellValue().trim();
	  String curr = sheet.getRow(i).getCell(6).getStringCellValue().trim();
	  String buyer = sheet.getRow(i).getCell(7).getStringCellValue().trim();
	  String line = sheet.getRow(i).getCell(8).getStringCellValue().trim();
	  String type = sheet.getRow(i).getCell(9).getStringCellValue().trim();
	  String item = sheet.getRow(i).getCell(10).getStringCellValue().trim();
	  String category = sheet.getRow(i).getCell(11).getStringCellValue().trim();
	  String requestor = sheet.getRow(i).getCell(12).getStringCellValue().trim();
	  String qnty = sheet.getRow(i).getCell(13).getStringCellValue().trim();
	  String price = sheet.getRow(i).getCell(14).getStringCellValue().trim();
	  String location = sheet.getRow(i).getCell(15).getStringCellValue().trim();
	  String rev = sheet.getRow(i).getCell(16).getStringCellValue().trim();
	  String reqDelDate = sheet.getRow(i).getCell(17).getStringCellValue().trim();
	  String proDelDate = sheet.getRow(i).getCell(18).getStringCellValue().trim();
	  
	  //** Purchase Order Creation **//
	  if(!procumentBU.equalsIgnoreCase("NA"))
	  {
	    
		 driver.findElement(By.xpath("//img[contains(@id, 'FndTasksList::icon')]")).click();
		 driver.findElement(By.linkText("Create Order")).click();
		 Thread.sleep(10000);
		 WebElement bussUnit = driver.findElement(By.xpath("//select[contains(@id, 'ProcurementBu::content')]"));
		 Select bussSel = new Select(bussUnit);
		 bussSel.selectByVisibleText(procumentBU);
		 Thread.sleep(8000);
		 WebElement bussUnit2 = driver.findElement(By.xpath("//select[contains(@id,'RequisitioningBu::content')]"));
		 Select bussSel2 = new Select(bussUnit2);
		 bussSel2.selectByVisibleText(requisionBU);
		 Thread.sleep(4000);
		 WebDriverWait wait = new WebDriverWait(driver, 350);
		 wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(@id, 'Supplier::lovIconId')]")));
		 driver.findElement(By.xpath("//a[contains(@id, 'Supplier::lovIconId')]")).click();
		 Thread.sleep(5000);
		 WebDriverWait supWait = new WebDriverWait(driver, 350);
		 supWait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[contains(@id, 'Supplier::_afrLovInternalQueryId:value00::content')]")));
		 driver.findElement(By.xpath("//input[contains(@id, 'Supplier::_afrLovInternalQueryId:value00::content')]")).sendKeys(supplier);
		 Thread.sleep(3000);
		 driver.findElement(By.xpath("//button[contains(@id, 'Supplier::_afrLovInternalQueryId::search')]")).click();
		 Thread.sleep(3000);
		 driver.findElement(By.xpath("//*[contains(@id,'Supplier_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		 driver.findElement(By.xpath("//button[contains(@id, 'Supplier::lovDialogId::ok')]")).click();
		 Thread.sleep(5000);
		 driver.findElement(By.xpath("//button[contains(@id, 'commandButton1')]")).click();
		 Thread.sleep(10000);
	  }
		 js.executeScript("window.scrollBy(0, 2000)", "");
		 driver.findElement(By.xpath("//img[contains(@id, 'AT1:_ATp:create::icon')]")).click();
		 Thread.sleep(8000);
		WebElement typ = driver.findElement(By.xpath("//input[contains(@id, 'LineType::content')]"));
		typ.click();
		typ.clear();
		typ.sendKeys(type);
		Thread.sleep(3000);
		typ.sendKeys(Keys.TAB);
		Thread.sleep(9000);
		driver.findElement(By.xpath("//a[contains(@id,'Item::lovIconId')]")).click();
		Thread.sleep(6000);
		driver.findElement(By.xpath("//input[contains(@id,'Id:value00::content')]")).sendKeys(item);
		Thread.sleep(2000);
		driver.findElement(By.xpath("//button[contains(@id,'QueryId::search')]")).click();
		Thread.sleep(4000);
		driver.findElement(By.xpath("//*[contains(@id,'LovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//button[contains(@id,'lovDialogId::ok')]")).click();
		Thread.sleep(33000);
		WebElement qn = driver.findElement(By.xpath("(//input[contains(@id, 'Quantity::content')])[1]"));
		//Thread.sleep(3000);
		JavascriptExecutor js11 = (JavascriptExecutor)driver;
		js11.executeScript("arguments[0].scrollIntoView();", qn);
		qn.click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("(//input[contains(@id, 'Quantity::content')])[1]")).sendKeys(qnty);
		driver.findElement(By.xpath("(//input[contains(@id, 'Quantity::content')])[1]")).sendKeys(Keys.TAB);
		Thread.sleep(7000);
		WebElement prc =driver.findElement(By.xpath("//input[contains(@id, 'UnitPrice::content')]"));
		JavascriptExecutor js3 = (JavascriptExecutor)driver;
		js3.executeScript("arguments[0].scrollIntoView();", prc);
		prc.click();
		prc.clear();
		prc.sendKeys(price);
//		Thread.sleep(5000);
		prc.sendKeys(Keys.TAB);
		Thread.sleep(5000);
		WebElement req = driver.findElement(By.xpath("//input[contains(@id, 'Requester::content')]"));
		JavascriptExecutor js4 = (JavascriptExecutor)driver;
		js4.executeScript("arguments[0].scrollIntoView();", req);
		req.click();
		req.sendKeys(requestor);
		Thread.sleep(4000);
		driver.findElement(By.xpath("//input[contains(@id, 'Requester::content')]")).sendKeys(Keys.TAB);
		Thread.sleep(5000);
		driver.findElement(By.xpath("(//a[text()='Schedules'])[1]")).click();
		Thread.sleep(8000);
		driver.findElement(By.xpath("//span[text()="+line+"]")).click();
		Thread.sleep(5000);
		driver.findElement(By.xpath("//input[contains(@id, 'NeedByDt::content')]")).click();
		driver.findElement(By.xpath("//input[contains(@id, 'NeedByDt::content')]")).sendKeys(reqDelDate);
		driver.findElement(By.xpath("//input[contains(@id, 'NeedByDt::content')]")).sendKeys(Keys.TAB);
		Thread.sleep(5000);
		driver.findElement(By.xpath("//input[contains(@id, 'PromisedDt::content')]")).click();
		driver.findElement(By.xpath("//input[contains(@id, 'PromisedDt::content')]")).sendKeys(proDelDate);
		driver.findElement(By.xpath("//input[contains(@id, 'PromisedDt::content')]")).sendKeys(Keys.TAB);
		Thread.sleep(8000);
		driver.findElement(By.xpath("(//a[text()='Lines'])[1]")).click();
		
		
	 }
	 }
	 else {
		 System.out.println("File is already Processed");
	 }
	 }
	 
	
	catch(Exception e)
	{
	e.printStackTrace();
	                 CellStyle style = wb.createCellStyle();
			  		 Font font = wb.createFont();
			  		 XSSFCell cell2 = sheet.getRow(i-1).createCell(20);
			  		 cell2.setCellValue("Fail");
			  		 font.setColor(IndexedColors.RED.getIndex());
			  		 font.setBold(true);
			  		 style = wb.createCellStyle();
			  		 style.setFont(font);
			  		 cell2.setCellStyle(style);
			  		 fos = new FileOutputStream(f);
			  		 wb.write(fos);
					 Assert.assertEquals(false, true);
	
	}
	
	finally
	{
	wb.close();
	fis.close();
	try {
	    if (fos != null)
	    {
	fos.close();
	    }
	         }
	 catch (IOException ioe) {
	    System.out.println("Error in closing the Stream");
	 }
	}  
	}
	    public void WebDriverwaitelement(WebElement element)
		{
			WebDriverWait wait = new WebDriverWait(driver,500);
			wait.until(ExpectedConditions.visibilityOf(element));
		}
	    public static boolean isRowEmpty(Row row) {
		    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
		        Cell cell = row.getCell(c);
		        if (cell != null && cell.getCellType() != CellType.BLANK)
		            return false;
		    }
		    return true;
		  }

}
