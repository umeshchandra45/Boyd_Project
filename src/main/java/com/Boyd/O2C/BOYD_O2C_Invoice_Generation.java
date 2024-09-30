package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class BOYD_O2C_Invoice_Generation {
	
	public static WebDriver driver;
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	 File f;
    XSSFSheet sheet;
    int i;
    int rowNumber;
	public String orderNumber;
	public String bussUnit;
	public String tranSource;
	
@Test
	
	public void orderStatusMgmt() throws InterruptedException, IOException
	{
		try {
			WebDriverManager.chromedriver().setup();
		    ChromeOptions options = new ChromeOptions();

		  driver = new ChromeDriver(options);
		  driver.manage().deleteAllCookies();
		  driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		  driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
		  driver.get("https://elme-dev1.fa.us8.oraclecloud.com/");
	//  driver.get("https://elme-dev2.fa.us8.oraclecloud.com");
//		  driver.get("https://elme-test.login.us8.oraclecloud.com/");
		  driver.manage().window().maximize();
		 
		  
//		  String remoteExecution = System.getProperty("remoteExecution");
//			System.out.println("remoteFlag  is :" +remoteExecution);
//			Boolean exectionFlag = Boolean.parseBoolean(remoteExecution);
			 File f = null;
//			if(exectionFlag) {
//				System.out.println("remote execution");
//				f = new File(System.getProperty("user.dir") + "\\ExternalFiles\\BOYD_O2C_InvoiceGeneration.xlsx");
//			}			 
//			else {
				 f = new File(System.getProperty("user.dir") + "\\Excel\\BOYD_O2C_InvoiceGeneration.xlsx");
//				 System.out.println("local execution");
//			}
		  fis = new FileInputStream(f);
		  wb = new XSSFWorkbook(fis);
		  sheet = wb.getSheet("InvoiceGeneration");
		  
		//browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
//			String url = System.getProperty("url");
//			System.out.println("url is :" +url);
//			String username = System.getProperty("userName");
//			System.out.println("username  is :" +username);
//			String password = System.getProperty("password");
//			System.out.println("password  is :" +password);
			
//			driver.get(url);
			driver.findElement(By.id("userid")).click();	
			//browser.findElement(By.id("userid")).sendKeys("forsys.user");
			
			
			driver.findElement(By.id("userid")).sendKeys("forsys.user");
			driver.findElement(By.id("password")).click();
			//browser.findElement(By.id("password")).sendKeys("welcome123");	
			driver.findElement(By.id("password")).sendKeys("forsys2023");
		  driver.findElement(By.id("btnActive")).click();
		  Thread.sleep(4000);
		  int rowNum=sheet.getPhysicalNumberOfRows();
		  System.out.println("rowNum="+rowNum);
		  System.out.println("colNum="+sheet.getRow(1).getLastCellNum());
		  Row row = sheet.getRow(1);
		  Cell c = row.getCell(3);
		  System.out.println("result=="+c);
		 if(c==null||c.getStringCellValue().contentEquals(""))
			 {
		  
		  for(i=1; i<rowNum; i++)
		  {
		
		  driver.findElement(By.id("pt1:_UIShome")).click();
		  Thread.sleep(10000);
		  driver.findElement(By.linkText("Order Management")).click();
		  Thread.sleep(3000);
		  driver.findElement(By.id("itemNode_order_management_order_management_1")).click();
		  orderNumber = sheet.getRow(i).getCell(0).getStringCellValue();
		  bussUnit = sheet.getRow(i).getCell(1).getStringCellValue();
		  tranSource = sheet.getRow(i).getCell(2).getStringCellValue();
		  Boolean flag=true;
		  search();
		  Thread.sleep(5000);
		  List<WebElement> ele = driver.findElements(By.xpath("//*[contains(@id, 'AP1:AT1:_ATp:ATt1::db')]/table/tbody/tr"));
		  int eleSize = ele.size();
		  System.out.println("Size of order table =="+eleSize);
		  if(eleSize==0)
		  {
			  System.out.println("Invalid Order Number");
			  CellStyle style = wb.createCellStyle();
	    		 Font font = wb.createFont();
	    		 XSSFCell cell2 = sheet.getRow(i).createCell(3);
	    		 cell2.setCellValue("Fail");
	    		 sheet.getRow(i).createCell(4).setCellValue("Invalid order number");
	    		 font.setColor(IndexedColors.RED.getIndex());
	    		 font.setBold(true);
	    		 style = wb.createCellStyle();
	    		 style.setFont(font);
	    		 cell2.setCellStyle(style);
	    		 fos = new FileOutputStream(f);
	    		 wb.write(fos);  
		  }
		  
		  else
		  {
		  driver.findElement(By.linkText(orderNumber)).click();
		  Thread.sleep(5000);
		  JavascriptExecutor js = (JavascriptExecutor) driver;
	       js.executeScript("window.scrollBy(0, 2000)", "");
		  List<WebElement> fisrttable = driver.findElements(By.xpath("//*[@id=\'pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:2:APVIEW1:pc1:ATt1::db\']/table/tbody/tr"));
		  int tablesize = fisrttable.size();
		  System.out.println("tablesize="+tablesize);
		  int j = 1;
		  
		  while(j<=tablesize)
		  {
		  String statusText=driver.findElement(By.xpath("//*[contains(@id, 'pc1:ATt1::db')]/table/tbody/tr["+j+"]/td[2]/div/table/tbody/tr/td[4]/span")).getText();
		  System.out.println("statusText="+statusText);
		  String actualStatus="Awaiting Billing";
		  Thread.sleep(2000);
		      if(statusText.equalsIgnoreCase(actualStatus))
		       {
		    	  Thread.sleep(5000);
		    	  driver.findElement(By.id("pt1:_UIShome::icon")).click();
		    	  Thread.sleep(2000);
		    	  js.executeScript("window.scrollBy(0, 2000)", "");
		    	  Thread.sleep(3000);
		    	  driver.findElement(By.id("groupNode_tools")).click();
		    	  driver.findElement(By.id("itemNode_tools_scheduled_processes_fuse_plus")).click();
		    	  Thread.sleep(3000);
		    	  driver.findElement(By.linkText("Schedule New Process")).click();
		    	  Thread.sleep(5000);
		    	  driver.findElement(By.xpath("//a[contains(@id, 'pt1:selectOneChoice2::lovIconId')]")).click();
		    	  js.executeScript("window.scrollBy(0, 1000)", "");
		    	  Thread.sleep(4000);
		    	  driver.findElement(By.xpath("//a[contains(@id, 'pt1:selectOneChoice2::dropdownPopup::popupsearch')]")).click();
		    	  Thread.sleep(5000);
		    	  driver.findElement(By.xpath("//input[contains(@id, 'pt1:selectOneChoice2::_afrLovInternalQueryId:value00::content')]")).sendKeys("Import AutoInvoice");
		    	  driver.findElement(By.xpath("//button[contains(@id, 'pt1:selectOneChoice2::_afrLovInternalQueryId::search')]")).click();
		    	  Thread.sleep(4000);
		    	  driver.findElement(By.xpath("//*[contains(@id, 'pt1:selectOneChoice2_afrLovInternalTableId::db')]/table/tbody/tr[1]/td[2]/div/table/tbody/tr/td[1]/span")).click();
		    	  driver.findElement(By.xpath("//button[contains(@id, 'pt1:selectOneChoice2::lovDialogId::ok')]")).click();
		    	  Thread.sleep(4000);
		    	  driver.findElement(By.xpath("//button[contains(@id, 'pt1:snpokbtnid')]")).click();
		    	  Thread.sleep(5000);
		    	  WebElement ele1 = driver.findElement(By.xpath("//input[contains(@id, 'basicReqBody:paramDynForm_BusinessUnit::content')]"));
		    	  ele1.clear();
		    	  ele1.sendKeys(bussUnit);
		    	  ele1.sendKeys(Keys.TAB);
		    	  Thread.sleep(5000);
		    	  driver.findElement(By.xpath("//input[contains(@id, 'basicReqBody:paramDynForm_ATTRIBUTE4_ATTRIBUTE4::content')]")).clear();
		    	  Thread.sleep(3000);
		    	  driver.findElement(By.xpath("//input[contains(@id, 'basicReqBody:paramDynForm_ATTRIBUTE4_ATTRIBUTE4::content')]")).sendKeys(tranSource);
		    	  Thread.sleep(5000);
		    	  WebElement ordernum = driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_ATTRIBUTE18_ATTRIBUTE18::content')]"));
		    	  ordernum.sendKeys(orderNumber);
		    	  ordernum.sendKeys(Keys.TAB);
//		    	  driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_ATTRIBUTE19_ATTRIBUTE19::content')]")).sendKeys(orderNumber);
		    	  Thread.sleep(8000);
		    	  driver.findElement(By.xpath("//div[contains(@id, 'requestBtns:submitButton')]")).click();
		    	  Thread.sleep(5000);
		    	  driver.findElement(By.xpath("//*[contains(@id, 'confirmationPopup:confirmSubmitDialog::ok')]")).click();
		    	  for(int m=0; m<=8;m++)
		    	  {   
		    	    Thread.sleep(6000);
		    	    driver.findElement(By.xpath("//img[contains(@id, 'pt1:panel:processRefreshId::icon')]")).click();
		    	  }
		    	  Thread.sleep(5000);
		    	  String statusVal=driver.findElement(By.xpath("//*[contains(@id,'panel:result::db')]/table/tbody/tr[1]/td[2]/div/table/tbody/tr/td[3]")).getText();
		  		  System.out.println("statusVal="+statusVal);
		  		  Thread.sleep(4000);
			  		if(!statusVal.equalsIgnoreCase("")) {
			  			
			  			  Thread.sleep(5000);
						  driver.findElement(By.id("pt1:_UIShome::icon")).click();
						  Thread.sleep(5000);
						  driver.findElement(By.id("groupNode_order_management")).click();
						  driver.findElement(By.id("itemNode_order_management_order_management")).click();
						  search();
						  Thread.sleep(40000);
						  driver.findElement(By.linkText(orderNumber)).click();
						  Thread.sleep(5000);
						  js.executeScript("window.scrollBy(0, 2000)", "");
						  Thread.sleep(5000);
						  String completeStatus=driver.findElement(By.xpath("//*[contains(@id, 'pc1:ATt1::db')]/table/tbody/tr[\"+j+\"]/td[2]/div/table/tbody/tr/td[4]/span")).getText();
						  System.out.println("completeStatus="+completeStatus);
						  String actStatus="Closed";
						  if(completeStatus.equalsIgnoreCase(actStatus))
						  {
							  flag=false;
							  System.out.println("Awaiting Billing Status is Closed");
							  CellStyle style = wb.createCellStyle();
			    	    		 Font font = wb.createFont();
			    	    		 XSSFCell cell2 = sheet.getRow(i).createCell(3);
			    	    		 cell2.setCellValue("Pass");
			    	    		 font.setColor(IndexedColors.GREEN.getIndex());
			    	    		 font.setBold(true);
			    	    		 style = wb.createCellStyle();
			    	    		 style.setFont(font);
			    	    		 cell2.setCellStyle(style);
			    	    		 fos = new FileOutputStream(f);
			    	    		 wb.write(fos);
			    	    		 break;
						  }
						  
						  else
						  {
							  flag=false;
							  System.out.println("Awaiting Billing Status is Not Closed");
							  CellStyle style = wb.createCellStyle();
			    	    		 Font font = wb.createFont();
			    	    		 XSSFCell cell2 = sheet.getRow(i).createCell(3);
			    	    		 cell2.setCellValue("Fail");
			    	    		 font.setColor(IndexedColors.RED.getIndex());
			    	    		 font.setBold(true);
			    	    		 style = wb.createCellStyle();
			    	    		 style.setFont(font);
			    	    		 cell2.setCellStyle(style);
			    	    		 sheet.getRow(i).createCell(4).setCellValue("Awaiting Billing status is not closed due to error");
			    	    		 fos = new FileOutputStream(f);
			    	    		 wb.write(fos); 
			    	    		 break;
						  }
						 
			  		}
		       }
		      else {
		    	  j++;
		    	  
		      }
		       }
		  if(flag)
		  {
		  System.out.println("Awaiting Billing status is not available in this order");
		  CellStyle style = wb.createCellStyle();
    		 Font font = wb.createFont();
    		 XSSFCell cell2 = sheet.getRow(i).createCell(3);
    		 cell2.setCellValue("Fail");
    		 font.setColor(IndexedColors.RED.getIndex());
    		 font.setBold(true);
    		 style = wb.createCellStyle();
    		 style.setFont(font);
    		 cell2.setCellStyle(style);
    		 sheet.getRow(i).createCell(4).setCellValue("Awaiting Billing status is not available");
    		 fos = new FileOutputStream(f);
    		 wb.write(fos); 
		  }
		  
		}}}
		 else
		 {
			 System.out.println("File is already processed");
		 }
		 
		 
		 
		 
	}
		catch(Exception e)
		{
			e.printStackTrace();
			CellStyle style = wb.createCellStyle();
    		 Font font = wb.createFont();
    		 XSSFCell cell2 = sheet.getRow(i).createCell(3);
    		 cell2.setCellValue("Fail");
    		 font.setColor(IndexedColors.RED.getIndex());
    		 font.setBold(true);
    		 style = wb.createCellStyle();
    		 style.setFont(font);
    		 cell2.setCellStyle(style);
    		 sheet.getRow(i).createCell(4).setCellValue("Exception Occured");
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
	public void search()
	{
		try {
		  driver.findElement(By.linkText("Advanced")).click();
		  Thread.sleep(4000);
		  WebElement ele2 = driver.findElement(By.xpath("//select[contains(@id, 'operator1')]"));
		  Select element11 = new Select(ele2);
		  element11.selectByVisibleText("Equals");
		  Thread.sleep(2000);
		  driver.findElement(By.xpath("//input[contains(@id, 'value10')]")).sendKeys(orderNumber);
		  WebElement state = driver.findElement(By.xpath("//select[contains(@id, 'operator7::content')]"));
		  Select element1 = new Select(state);
		  element1.selectByVisibleText("Equals");
		  Thread.sleep(2000);
		  WebElement process = driver.findElement(By.xpath("//select[contains(@id, 'value70::content')]"));
		  Select element2 = new Select(process);
		  element2.selectByVisibleText("Processing");
		  Thread.sleep(3000);
		  driver.findElement(By.xpath("//button[contains(@id, 'search')]")).click();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

}
