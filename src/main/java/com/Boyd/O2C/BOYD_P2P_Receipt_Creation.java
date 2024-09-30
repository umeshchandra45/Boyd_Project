package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class BOYD_P2P_Receipt_Creation {
	
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	    WebDriver driver;
	    String result;
	   //int i;
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
	 driver.get("https://elme-dev1.fa.us8.oraclecloud.com");
	 driver.manage().window().maximize();

	 
//	 String remoteExecution = System.getProperty("remoteExecution");
//		System.out.println("remoteFlag  is :" +remoteExecution);
//		Boolean exectionFlag = Boolean.parseBoolean(remoteExecution);
		File f = null;
//		if(exectionFlag) {
//			System.out.println("remote execution");
//			f = new File(System.getProperty("user.dir") + "\\ExternalFiles\\BOYD_P2P_Reciept_Creation.xlsx");
//		}			 
//		else {
			 f = new File(System.getProperty("user.dir") + "\\Excel\\BOYD_P2P_Reciept_Creation.xlsx");
//			 System.out.println("local execution");
//		}
	 
	 fis = new FileInputStream(f);
	 wb = new XSSFWorkbook(fis);
	 XSSFSheet sheet = wb.getSheet("Receipt Creation");
	 
	//browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
//			String url = System.getProperty("url");
//			System.out.println("url is :" +url);
//			String username = System.getProperty("userName");
//			System.out.println("username  is :" +username);
//			String password = System.getProperty("password");
//			System.out.println("password  is :" +password);
			
//			driver.get("https://elme-dev2.fa.us8.oraclecloud.com");
//	 driver.get("https://elme-test.login.us8.oraclecloud.com/");
			driver.findElement(By.id("userid")).click();	
			//browser.findElement(By.id("userid")).sendKeys("forsys.user");
			
			
			driver.findElement(By.id("userid")).sendKeys("forsys.user");
			driver.findElement(By.id("password")).click();
			driver.findElement(By.id("password")).sendKeys("forsys2023");	
	//		driver.findElement(By.id("password")).sendKeys("forsys4@4!");
	 driver.findElement(By.id("btnActive")).click();
	 Thread.sleep(5000);
	  WebDriverWait wait1 = new WebDriverWait(driver, 500);
	  wait1.until(ExpectedConditions.elementToBeClickable(By.id("pt1:_UIShome")));
	  driver.findElement(By.id("pt1:_UIShome")).click();
	 Thread.sleep(15000);
	 driver.findElement(By.linkText("Supply Chain Execution")).click();
	 driver.findElement(By.linkText("Inventory Management")).click();
	 Thread.sleep(12000);
	 JavascriptExecutor js = (JavascriptExecutor) driver;
	 int rowNum=sheet.getPhysicalNumberOfRows();
	 System.out.println("rowNum="+rowNum);
	 System.out.println("colNum="+sheet.getRow(1).getLastCellNum());
	 Row row = sheet.getRow(1);
	  Cell c = row.getCell(5);
	  System.out.println("result=="+c);
	  
	  
	 if(c==null||c.getStringCellValue().contentEquals(""))
		 {
		//**Start Receipt Creation **//
	 for(int i=1; i<=rowNum;)
	 {
	 
		 if(sheet.getRow(i) == null || isRowEmpty(sheet.getRow(i))) {
			 try
			 {
		    driver.findElement(By.xpath("//button[contains(@id, 'appPanelid:cb3')]")).click();
	        Thread.sleep(6000);
	        try
	        {
	        	driver.findElement(By.xpath("//button[contains(@id, 'msgDlg::cancel')]")).click();
	        	Thread.sleep(2000);
	        	driver.findElement(By.xpath("//button[contains(@id, 'appPanelid:cb3')]")).click();
		        Thread.sleep(6000);
	        }
	        catch(Exception e)
	        {
	        	
	        }
            driver.findElement(By.xpath("//*[contains(@id, 'ap1:SPsb2')]/a/span")).click();
	    	  WebDriverWait waitt = new WebDriverWait(driver, 350);
			  waitt.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id, 'ap1:confDlg::contentContainer')]/span")));
	    	  String confmText = driver.findElement(By.xpath("//*[contains(@id, 'ap1:confDlg::contentContainer')]/span")).getText();
	    	  //Thread.sleep(5000);
	    	  System.out.println("confirmText="+confmText);
	    	  Thread.sleep(3000);
	  		String[] str3=confmText.split(" ");
	  		Thread.sleep(3000);
	  		String str2 = str3[1];
            System.out.println("recNum="+str2);
	    	XSSFCell cell11 = sheet.getRow(i-1).createCell(4);
	    	 cell11.setCellValue(str2);
	    	 fos = new FileOutputStream(f);
	    	 wb.write(fos);
	    	 Thread.sleep(5000);
	    	 CellStyle style = wb.createCellStyle();
	    	 Font font = wb.createFont();
	    	 XSSFCell cell2 = sheet.getRow(i-1).createCell(5);
	    	    cell2.setCellValue("Pass");
	    	    font.setColor(IndexedColors.GREEN.getIndex());
	    		font.setBold(true);
	    		style = wb.createCellStyle();
	    		style.setFont(font);
	    		cell2.setCellStyle(style);
	    	    fos = new FileOutputStream(f);
	    	    wb.write(fos);
	    	    driver.findElement(By.xpath("//span[text()='K']")).click();
	    	    Thread.sleep(4000);
	    	    driver.findElement(By.xpath("//*[contains(@id, 'pt1:ap1:SPb')]")).click();  
	    	    i++;
			 }
			 catch(Exception e)
			 {
				 i++;
			 }
	    	    continue;
		 }
	  String PONum = sheet.getRow(i).getCell(0).getStringCellValue().trim();
	  
	 
	  
		//**Inventory**//
	  if(!PONum.equalsIgnoreCase("NA"))
	  {
		 driver.findElement(By.xpath("//img[contains(@id, 'FOTsdiScmInvOverviewPage_itemNode_InvTasksList::icon')]")).click();
		 Thread.sleep(6000);
	   Select receipts = new Select(driver.findElement(By.xpath("//select[contains(@id, 'FOTRaT:0:soc1::content')]")));
	   Thread.sleep(5000);
	   receipts.selectByVisibleText("Receipts");
	   Thread.sleep(4000);
		  driver.findElement(By.linkText("Receive Expected Shipments")).click();
		  Thread.sleep(5000);
		  driver.findElement(By.xpath("//input[contains(@id, 'rcvQry:value00::content')]")).sendKeys(PONum);
		  Thread.sleep(4000);
		  driver.findElement(By.xpath("//button[contains(@id, 'ap1:rcvQry::search')]")).click();
		  Thread.sleep(4000);
	  
		  List<WebElement> table = driver.findElements(By.xpath("//*[contains(@id, 'ap1:AT1:_ATp:rcv::db')]/table/tbody/tr"));
	    int tablesize = table.size();
	    System.out.println("table size="+tablesize);
	    if(tablesize<=0)
	    {
	  	  System.out.println("There is no data found in Purchase Order");
	  	  CellStyle style = wb.createCellStyle();
	  		 Font font = wb.createFont();
	  		 XSSFCell cell2 = sheet.getRow(i).createCell(5);
	  		 cell2.setCellValue("Fail");
	  		 font.setColor(IndexedColors.RED.getIndex());
	  		 font.setBold(true);
	  		 style = wb.createCellStyle();
	  		 style.setFont(font);
	  		 cell2.setCellStyle(style);
	  		sheet.getRow(i).createCell(6).setCellValue("No data found with the given PO number");
	  		 fos = new FileOutputStream(f);
	  		 wb.write(fos);
	  		driver.findElement(By.xpath("//*[contains(@id, 'pt1:ap1:SPb')]")).click();
	  		 i++;
	    }
	    else
	    {

	        for(int k=1; k<=tablesize; k++)
	         {
	            WebElement Quantity = driver.findElement(By.xpath("//*[contains(@id, 'ap1:AT1:_ATp:rcv::db')]/table/tbody/tr["+k+"]/td[2]/div/table/tbody/tr/td[9]/span"));
	            String Qnty1 = Quantity.getText();
	            System.out.println("Quantity="+Qnty1); 
	            
	            if(tablesize==1)
	            {
	            	driver.findElement(By.xpath("//*[contains(@id, 'ap1:AT1:_ATp:rcv::db')]/table/tbody/tr/td[1]")).click();
	            }
	            else
	            {
	            Actions builder = new Actions(driver);
	            builder.click(table.get(0)).keyDown(Keys.CONTROL).click(table.get(tablesize-1)).keyUp(Keys.CONTROL).build().perform();
	            }
	         }
	          Thread.sleep(4000);
	    	  driver.findElement(By.xpath("//*[contains(@id, 'ap1:AT1:_ATp:receive')]")).click();
	    	  Thread.sleep(6000);
	    
	    	  List<WebElement> table1 = driver.findElements(By.xpath("//*[contains(@id, 'appPanelid:AT1:_ATp:table1::db')]/table/tbody/tr"));
	  	    int tablesize1 = table1.size();
	  	    System.out.println("table size1=="+tablesize1);
	    
	  	  for(int k=1; k<=tablesize1; k++)
	         {
	  		String lineNum = sheet.getRow(i).getCell(1).getStringCellValue().trim();
	  	    String item = sheet.getRow(i).getCell(2).getStringCellValue().trim();
	  	    String qnty = sheet.getRow(i).getCell(3).getStringCellValue().trim();
	  		driver.findElement(By.xpath("//span[text()='"+item+"']/../../../..//input[contains(@id,'Quantityid::content')]")).sendKeys(qnty);
	  		Thread.sleep(4000);
	  		i++;
	         }
	    	
	    	    
	    	   
	    }
	      }
	  else
		  {
		  i++;
		  }
		  }
	//**End Receipt Creation **//
	 
	 //**Start Inspection **//
	 Thread.sleep(8000);
	 XSSFSheet sheet1 = wb.getSheet("Inspection");
	 int rowNum1=sheet1.getPhysicalNumberOfRows();
	 System.out.println("rowNum1=="+rowNum1);
	 
	 for(int j=1; j<=rowNum1;)
	 {
		 if(sheet1.getRow(j) == null || isRowEmpty(sheet1.getRow(j))) {
			 try
			 {
				 driver.findElement(By.xpath("//div[contains(@id, 'SPsb2')]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//button[contains(@id, 'cb2')]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//div[contains(@id, 'SPb')]")).click();
				 Thread.sleep(5000);
		    	 CellStyle style = wb.createCellStyle();
		    	 Font font = wb.createFont();
		    	 XSSFCell cell2 = sheet1.getRow(j-1).createCell(11);
		    	    cell2.setCellValue("Pass");
		    	    font.setColor(IndexedColors.GREEN.getIndex());
		    		font.setBold(true);
		    		style = wb.createCellStyle();
		    		style.setFont(font);
		    		cell2.setCellStyle(style);
		    	    fos = new FileOutputStream(f);
		    	    wb.write(fos);
		    	    j++;
				  
			 }
			 catch(Exception e)
			 {
				 j++;
			 }
	    	    continue;
		 }
	  String PONum1 = sheet1.getRow(j).getCell(0).getStringCellValue().trim();
	  String receiptNum = sheet1.getRow(j).getCell(1).getStringCellValue().trim();
	  
	  if(!PONum1.equalsIgnoreCase("NA"))
		  {
		 driver.findElement(By.xpath("//img[contains(@id, 'FOTsdiScmInvOverviewPage_itemNode_InvTasksList::icon')]")).click();
		 Thread.sleep(3000);
	     driver.findElement(By.xpath("//select[contains(@id, 'FOTRaT:0:soc1::content')]")).sendKeys("Receipts");
		 driver.findElement(By.linkText("Inspect Receipts")).click();
		 Thread.sleep(5000);
		 driver.findElement(By.xpath("//input[contains(@id, 'insQry:value10::content')]")).sendKeys(PONum1);
		 driver.findElement(By.xpath("//button[contains(@id, 'insQry::search')]")).click();
		 Thread.sleep(5000);
		 List<WebElement> insTable = driver.findElements(By.xpath("//*[contains(@id, 'ins::db')]/table/tbody/tr"));
		    int insTableSize = insTable.size();
		    System.out.println("Inspection table size="+insTableSize);
		    if(insTableSize<=0)
			    {
			  	  System.out.println("There is no data found in Purchase Order");
			  	  CellStyle style = wb.createCellStyle();
			  		 Font font = wb.createFont();
			  		 XSSFCell cell2 = sheet1.getRow(j).createCell(11);
			  		 cell2.setCellValue("Fail");
			  		 font.setColor(IndexedColors.RED.getIndex());
			  		 font.setBold(true);
			  		 style = wb.createCellStyle();
			  		 style.setFont(font);
			  		 cell2.setCellStyle(style);
			  		sheet1.getRow(j).createCell(12).setCellValue("No data found with the given PO number");
			  		 fos = new FileOutputStream(f);
			  		 wb.write(fos);
			  		driver.findElement(By.xpath("//*[contains(@id, 'pt1:ap1:SPb')]")).click();
			  		 j++;
			    }
		    else
		    {
		    	for(int k=1; k<=insTableSize; k++)
			         {  
		    		if(insTableSize==1)
		            {
		            	driver.findElement(By.xpath("//*[contains(@id, 'ins::db')]/table/tbody/tr/td[1]")).click();
		            }
		    		else {
			            Actions builder = new Actions(driver);
			            builder.click(insTable.get(0)).keyDown(Keys.CONTROL).click(insTable.get(insTableSize-1)).keyUp(Keys.CONTROL).build().perform();
		    		}
			         }
		    
		    	Thread.sleep(4000);
		 driver.findElement(By.xpath("//button[contains(@id, 'inspect')]")).click();
		 Thread.sleep(6000);
		 List<WebElement> insTable1 = driver.findElements(By.xpath("//*[contains(@id, 'table1::db')]/table/tbody/tr"));
		    int insTableSize1 = insTable1.size();
		    System.out.println("Inspection table size1="+insTableSize1);
		    for(int m=1; m<=insTableSize1; m++)
		         {
		      String lineNum1 = sheet1.getRow(j).getCell(2).getStringCellValue().trim();
		  	  String item1 = sheet1.getRow(j).getCell(3).getStringCellValue().trim();
		  	  String insResult1 = sheet1.getRow(j).getCell(5).getStringCellValue().trim();
		  	  String insResult2 = sheet1.getRow(j).getCell(6).getStringCellValue().trim();
		  	  String insResult3 = sheet1.getRow(j).getCell(7).getStringCellValue().trim();
		  	  String insResult4 = sheet1.getRow(j).getCell(8).getStringCellValue().trim();
		  	  String insResult5 = sheet1.getRow(j).getCell(9).getStringCellValue().trim();
		  		driver.findElement(By.xpath("//*[text()= '"+item1+"']")).click();
		  		Thread.sleep(4000);
		  		driver.findElement(By.xpath("//button[contains(@id, 'eqrb')]")).click();
				 Thread.sleep(5000);
				 List<WebElement> quaTable = driver.findElements(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr"));
				    int quaTableSize = quaTable.size();
				    System.out.println("Quality table size1="+quaTableSize);
				    if(quaTableSize==5)
				    {
				 driver.findElement(By.xpath("//input[contains(@id, '0:it51::content')]")).sendKeys(insResult1);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '0:it51::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[2]/td[1]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '1:it51::content')]")).sendKeys(insResult2);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '1:it51::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[3]/td[1]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '2:iclov2::content')]")).sendKeys(insResult3);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '2:iclov2::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[4]/td[1]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '3:iclov2::content')]")).sendKeys(insResult4);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '3:iclov2::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 WebElement ele = driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[4]/td[1]"));
				 js.executeScript("arguments[0].scrollIntoView();",ele );
				 ele.click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '4:iclov2::content')]")).sendKeys(insResult5);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '4:iclov2::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				    }
				    
				    if(quaTableSize==4)
				    {
				 driver.findElement(By.xpath("//input[contains(@id, '0:it51::content')]")).sendKeys(insResult1);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '0:it51::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[2]/td[1]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '1:it51::content')]")).sendKeys(insResult2);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '1:it51::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[3]/td[1]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '2:iclov2::content')]")).sendKeys(insResult3);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '2:iclov2::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[4]/td[1]")).click();
				 Thread.sleep(5000);
//				 driver.findElement(By.xpath("//input[contains(@id, 'iclov2::content')]")).sendKeys(insResult4);
//				 Thread.sleep(2000);
//				 driver.findElement(By.xpath("//input[contains(@id, 'iclov2::content')]")).sendKeys(Keys.TAB);
//				 Thread.sleep(5000);
//				 WebElement ele = driver.findElement(By.xpath("//*[contains(@id, 'AP1:AT3:_ATp:ATt3::db')]/table/tbody/tr[5]/td[1]"));
//				 js.executeScript("arguments[0].scrollIntoView();",ele );
//				 ele.click();
//				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//input[contains(@id, '3:iclov2::content')]")).sendKeys(insResult5);
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[contains(@id, '3:iclov2::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
				    }
				 
				 driver.findElement(By.xpath("//button[contains(@id, 'AP1:cb4')]")).click();
				 Thread.sleep(5000);
		  		j++;
		         }}
		  }
	  else
		  {
		  j++;
		  }
	 }
	 
	//**End Inspection **//
	 
	 //**Start Put Away **//
	 Thread.sleep(7000);
	 XSSFSheet sheet2 = wb.getSheet("Put Away Receipt");
	 int rowNum2=sheet2.getPhysicalNumberOfRows();
	 System.out.println("rowNum2=="+rowNum2);
	 
	 for(int n=1; n<=rowNum2;)
	 {
	 
		 if(sheet2.getRow(n) == null || isRowEmpty(sheet2.getRow(n))) {
			 try
			 {
				 driver.findElement(By.xpath("//div[contains(@id, 'ap1:SPsb2')]")).click();
				 WebDriverWait cnfwait = new WebDriverWait(driver, 350);
				 cnfwait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(@id, 'appTb:cnfbtn')]")));
				 driver.findElement(By.xpath("//button[contains(@id, 'appTb:cnfbtn')]")).click();
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//div[contains(@id, 'SPb')]")).click();
				 Thread.sleep(5000);
		    	 CellStyle style = wb.createCellStyle();
		    	 Font font = wb.createFont();
		    	 XSSFCell cell2 = sheet2.getRow(n-1).createCell(10);
		    	    cell2.setCellValue("Pass");
		    	    font.setColor(IndexedColors.GREEN.getIndex());
		    		font.setBold(true);
		    		style = wb.createCellStyle();
		    		style.setFont(font);
		    		cell2.setCellStyle(style);
		    	    fos = new FileOutputStream(f);
		    	    wb.write(fos);
				  n++;
			 }
			 catch(Exception e)
			 {
				 n++;
			 }
	    	    continue;
		 }
	  String PONum2 = sheet2.getRow(n).getCell(0).getStringCellValue().trim();
	  String receiptNum1 = sheet2.getRow(n).getCell(1).getStringCellValue().trim();
	 
	  if(!PONum2.equalsIgnoreCase("NA"))
		  {
		 driver.findElement(By.xpath("//img[contains(@id, 'FOTsdiScmInvOverviewPage_itemNode_InvTasksList::icon')]")).click();
		 Thread.sleep(3000);
	     driver.findElement(By.xpath("//select[contains(@id, 'FOTRaT:0:soc1::content')]")).sendKeys("Receipts");
		 driver.findElement(By.linkText("Put Away Receipts")).click();
		 Thread.sleep(5000);
		 driver.findElement(By.xpath("//input[contains(@id, 'delQry:value10::content')]")).sendKeys(PONum2);
		 driver.findElement(By.xpath("//button[contains(@id, 'delQry::search')]")).click();
		 Thread.sleep(5000);
		 List<WebElement> putAwayTable = driver.findElements(By.xpath("//*[contains(@id, 'put::db')]/table/tbody/tr"));
		    int putTableSize = putAwayTable.size();
		    System.out.println("PytAway table size="+putTableSize);
		    if(putTableSize<=0)
		    {
		  	  System.out.println("There is no data found in Purchase Order");
		  	  CellStyle style = wb.createCellStyle();
		  		 Font font = wb.createFont();
		  		 XSSFCell cell2 = sheet2.getRow(n).createCell(10);
		  		 cell2.setCellValue("Fail");
		  		 font.setColor(IndexedColors.RED.getIndex());
		  		 font.setBold(true);
		  		 style = wb.createCellStyle();
		  		 style.setFont(font);
		  		 cell2.setCellStyle(style);
		  		sheet2.getRow(n).createCell(11).setCellValue("No data found with the given PO number");
		  		 fos = new FileOutputStream(f);
		  		 wb.write(fos);
		  		driver.findElement(By.xpath("//*[contains(@id, 'pt1:ap1:SPb')]")).click();
		  		 n++;
		    }
		    else
		    {
	
		        for(int k=1; k<=putTableSize; k++)
		         {  
		        	if(putTableSize==1)
		            {
		            	driver.findElement(By.xpath("//*[contains(@id, 'put::db')]/table/tbody/tr/td[1]")).click();
		            }
		        	else {
		            Actions builder = new Actions(driver);
		            builder.click(putAwayTable.get(0)).keyDown(Keys.CONTROL).click(putAwayTable.get(putTableSize-1)).keyUp(Keys.CONTROL).build().perform();
		        	}
		         }
		        Thread.sleep(4000);
		        driver.findElement(By.xpath("//button[contains(@id, 'deliver')]")).click();
		    	  Thread.sleep(6000);
		    
		    List<WebElement> putAwayTable1 = driver.findElements(By.xpath("//*[contains(@id, 'txTbl::db')]/table/tbody/tr"));
		    int putTableSize1 = putAwayTable1.size();
		    System.out.println("PutAway table size1="+putTableSize1);
		    for(int p=1; p<=putTableSize1; p++)
	         {
		      String lineNum2 = sheet2.getRow(n).getCell(2).getStringCellValue().trim();
		   	  String item2 = sheet2.getRow(n).getCell(3).getStringCellValue().trim();
		   	  String qantity = sheet2.getRow(n).getCell(4).getStringCellValue().trim();
		   	  String subInventory = sheet2.getRow(n).getCell(5).getStringCellValue().trim();
		   	  String locator = sheet2.getRow(n).getCell(6).getStringCellValue().trim();
		   	  String lot = sheet2.getRow(n).getCell(7).getStringCellValue().trim();
		   	  String serial = sheet2.getRow(n).getCell(8).getStringCellValue().trim();
		   	  
		   	  WebElement itm =driver.findElement(By.xpath("//*[text()= '"+item2+"']"));
		   	js.executeScript("arguments[0].scrollIntoView();",itm );
		   	Thread.sleep(3000);
		   	driver.findElement(By.xpath("//*[text()= '"+item2+"']")).click();
	  		Thread.sleep(4000);
	  		WebElement subInv = driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'sinv::content')]"));
			 js.executeScript("arguments[0].scrollIntoView();",subInv );
	  		driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'sinv::content')]")).sendKeys(subInventory);
			 driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'sinv::content')]")).sendKeys(Keys.TAB);
			 Thread.sleep(5000);
			 if(!locator.equalsIgnoreCase("NA"))
			 {
				 driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'kf1CS::content')]")).sendKeys(locator);
				 driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'kf1CS::content')]")).sendKeys(Keys.TAB);
			 }
			 
			 Thread.sleep(5000);
			 if(!lot.equalsIgnoreCase("NA"))
			 {
				 driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'lotText::content')]")).sendKeys(lot);
//				 driver.findElement(By.xpath("//*[text()= '"+item2+"']/../../..//input[contains(@id, 'lotText::content')]")).sendKeys(Keys.TAB);
				 Thread.sleep(5000);
			 }
			 n++;
	         }}}
		    else
				  {
				  n++;
				  }
		 
		 
	 }
	//**End Put Away **//
	 }

	 }
	
	catch(Exception e)
	{
	e.printStackTrace();
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
