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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import org.testng.Assert;
import io.github.bonigarcia.wdm.WebDriverManager;

@SuppressWarnings("unused")
public class BOYD_O2C_Material_Shipping {
	
	static FileInputStream fis;
	static FileOutputStream fos;
	static XSSFWorkbook wb;
	static WebDriver driver;
	int y=0;
   
    @Test
	public void partialShipment() throws InterruptedException, IOException
	{
		try {
			WebDriverManager.chromedriver().setup();
		  driver = new ChromeDriver();
		  driver.manage().deleteAllCookies();
		  driver.manage().timeouts().implicitlyWait(80, TimeUnit.SECONDS);
		  driver.manage().timeouts().pageLoadTimeout(80, TimeUnit.SECONDS);
		  driver.get("https://elme-dev1.fa.us8.oraclecloud.com/");
	//    driver.get("https://elme-dev2.fa.us8.oraclecloud.com");
//		  driver.get("https://elme-test.login.us8.oraclecloud.com/");
		  driver.manage().window().maximize();
		  File srcFile = new File(System.getProperty("user.dir")+"\\Excel\\BOYD_O2C_MaterialShipping.xlsx");
		  fis = new FileInputStream(srcFile);
		  wb = new XSSFWorkbook(fis);
		  XSSFSheet sheet = wb.getSheet("MaterialShipping");
		  driver.findElement(By.id("userid")).sendKeys("forsys.user");
   		  driver.findElement(By.id("password")).sendKeys("forsys2023");
		  driver.findElement(By.id("btnActive")).click();
		  Thread.sleep(5000);
		  WebDriverWait wait1 = new WebDriverWait(driver, 500);
		  wait1.until(ExpectedConditions.elementToBeClickable(By.id("pt1:_UIShome")));
		  driver.findElement(By.id("pt1:_UIShome")).click();
		  Thread.sleep(10000);
		  WebDriverWait wait = new WebDriverWait(driver, 350);
		  wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Supply Chain Execution")));
		  driver.findElement(By.linkText("Supply Chain Execution")).click();
		  driver.findElement(By.linkText("Inventory Management")).click();
		  int rowNum=sheet.getPhysicalNumberOfRows();
		  System.out.println("rowNum="+rowNum);
		  System.out.println("colNum="+sheet.getRow(1).getLastCellNum());
		  JavascriptExecutor js = (JavascriptExecutor) driver;
		  Row row = sheet.getRow(1);
		  Cell c = row.getCell(13);
		  System.out.println("result=="+c);
		 if(c==null||c.getStringCellValue().contentEquals(""))
			 {
		  
	  
		 //** Manage Shipment Lines**//
		  for(int rowNumber=1; rowNumber<rowNum;)
			  {
			  if(sheet.getRow(rowNumber) == null || isRowEmpty(sheet.getRow(rowNumber))) {
					 
			        Thread.sleep(3000);
				    rowNumber++;
				    continue;
		       }  
				  String orderNumber = sheet.getRow(rowNumber).getCell(0).getStringCellValue().trim();
				  String shipMethod = sheet.getRow(rowNumber).getCell(2).getStringCellValue().trim();
				  String warehouse = sheet.getRow(rowNumber).getCell(3).getStringCellValue().trim();
				  String wayBill = sheet.getRow(rowNumber).getCell(4).getStringCellValue().trim();
				  String weight = sheet.getRow(rowNumber).getCell(5).getStringCellValue().trim();
				  String volume = sheet.getRow(rowNumber).getCell(6).getStringCellValue().trim();
				  
				 
				  int shipmentLink = 0;
				  
				 if(!orderNumber.equalsIgnoreCase("NA"))
				 {
			 Thread.sleep(3000);
			 driver.findElement(By.xpath("//img[contains(@id, 'InvTasksList::icon')]")).click();
			 Thread.sleep(5000);
			 driver.findElement(By.xpath("//*[contains(@id, 'soc1::content')]")).sendKeys("Shipments");
			 driver.findElement(By.linkText("Manage Shipment Lines")).click();
			 Thread.sleep(3000);
			 //driver.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_warehouse_operations_InventoryManagement1:0:MAnt2:4:pt1:AP1:q1:value00::content")).sendKeys(orderNumber);
			 
			 WebElement order = driver.findElement(By.xpath("//*[contains(@id, 'AP1:q1:value00::content')]"));
			 order.click();
			 order.clear();
			 order.sendKeys(orderNumber);
			 driver.findElement(By.xpath("//*[contains(@id, 'AP1:q1:operator1::content')]")).sendKeys("After");
			 Thread.sleep(3000);
			 driver.findElement(By.xpath("//input[contains(@id,'pt1:AP1:q1:value10::content')]")).click();
			 driver.findElement(By.xpath("//input[contains(@id,'pt1:AP1:q1:value10::content')]")).sendKeys("01/03/24 12:45 PM");
			 Thread.sleep(8000);
//			 WebElement ele = driver.findElement(By.xpath("//select[contains(@id, 'AP1:q1:value50::content')]"));
//			 Select sel = new Select(ele);
//			 sel.selectByIndex(0);
//			 driver.findElement(By.xpath("//*[contains(@id, 'AP1:q1:value60::content')]")).sendKeys(warehouse);
			 driver.findElement(By.xpath("//*[contains(@id, 'AP1:q1::search')]")).click();
			 Thread.sleep(6000);
			 try
			 {
				 boolean isPresent = driver.findElements(By.xpath("//div[contains(@title,'Edit Shipment Line')]")).size()>0;
				 System.out.println("isPresent="+isPresent);
				 if(isPresent)
				 {
				  Thread.sleep(5000);
				 driver.findElement(By.xpath("//span[text()='ancel']")).click();
				 Thread.sleep(5000);
			 }
			 
			  }
			 catch(Exception ex)
			 {
				 
			 }
			 js.executeScript("window.scrollBy(0, 2000)", "");
			 Thread.sleep(3000);
			 List<WebElement> fisrttable = driver.findElements(By.xpath("//*[contains(@id, 'AP1:r12345:0:scat1:_ATp:table1::db')]/table/tbody/tr"));
			 int tablesize = fisrttable.size();
			 System.out.println("tablesize="+tablesize);
			 if(tablesize==0)
			 {
				 System.out.println("No data found with the given order number");
				  CellStyle style = wb.createCellStyle();
					Font font = wb.createFont();
				    XSSFCell cell2 = sheet.getRow(rowNumber).createCell(13);
				    cell2.setCellValue("Fail");
				    font.setColor(IndexedColors.RED.getIndex());
					font.setBold(true);
					style = wb.createCellStyle();
					style.setFont(font);
					cell2.setCellStyle(style);
				    XSSFCell cell = sheet.getRow(rowNumber).createCell(14);
				    cell.setCellValue("No data found with the given order number");
				    fos = new FileOutputStream(srcFile);
				    wb.write(fos);
				    rowNumber++;
			 }
			 else
			 {
			 while(shipmentLink<tablesize)
			 {
			 String lineStatus = driver.findElement(By.xpath("//*[contains(@id, 'AP1:r12345:0:scat1:_ATp:table1::db')]/table/tbody/tr["+(shipmentLink+1)+"]/td[4]/div/table/tbody/tr/td[9]/span")).getText();
			 System.out.println("lineStatus=="+lineStatus);
			 if(lineStatus.equalsIgnoreCase("Staged"))
			 {
			 String shipmentlk = driver.findElement(By.xpath("//*[contains(@id, 'AP1:r12345:0:scat1:_ATp:table1::db')]/table/tbody/tr["+(shipmentLink+1)+"]/td[3]")).getText();
			 System.out.println("shipmentlk="+shipmentlk);
			 Thread.sleep(5000); 
			 driver.findElement(By.linkText(shipmentlk)).click();
			 Thread.sleep(10000);
			 WebElement shipMethod1 = driver.findElement(By.xpath("//input[contains(@id, 'inputComboboxListOfValues1::content')]"));
			 shipMethod1.clear();
			 shipMethod1.sendKeys(shipMethod);
			 shipMethod1.sendKeys(Keys.TAB);
			 Thread.sleep(4000);  
			 driver.findElement(By.xpath("//input[contains(@id, 'inputText39::content')]")).clear();
			 driver.findElement(By.xpath("//input[contains(@id, 'inputText39::content')]")).sendKeys(wayBill);
			 Thread.sleep(3000);
			 driver.findElement(By.xpath("//input[contains(@id, 'pt1:scap1:inputText188::content')]")).clear();
			 driver.findElement(By.xpath("//input[contains(@id, 'pt1:scap1:inputText188::content')]")).sendKeys(weight);
			 driver.findElement(By.xpath("//input[contains(@id, 'pt1:scap1:inputText4::content')]")).clear();
			 driver.findElement(By.xpath("//input[contains(@id, 'pt1:scap1:inputText4::content')]")).sendKeys(volume);
		     js.executeScript("window.scrollBy(0, 2000)", "");
		     List<WebElement> seconTable = driver.findElements(By.xpath("//*[contains(@id, 'scat1:_ATp:table1::db')]/table/tbody/tr"));
			 int secTabSize = seconTable.size();
			 System.out.println("secTabSize="+secTabSize);
			 
			 for(int k=1;k<=secTabSize;k++)
			 {
				 String item = sheet.getRow(rowNumber).getCell(1).getStringCellValue().trim();
				 int lineNumber = (int) sheet.getRow(rowNumber).getCell(7).getNumericCellValue();
				  String line = String.valueOf(lineNumber).trim();
				  int reqQnty = (int) sheet.getRow(rowNumber).getCell(8).getNumericCellValue();
				  String rQnty = String.valueOf(reqQnty).trim();
				  int shipQnty = (int) sheet.getRow(rowNumber).getCell(9).getNumericCellValue();
				  String sQnty = String.valueOf(shipQnty).trim();
				  int trackNum = (int) sheet.getRow(rowNumber).getCell(10).getNumericCellValue();
				  String tNumb = String.valueOf(trackNum).trim();
				  
				 driver.findElement(By.xpath("//*[text()= '"+item+"']")).click();
				 Thread.sleep(4000);
				 WebElement shipmentQnty= driver.findElement(By.xpath("//input[contains(@id,'shpQtyInput::content')]"));
					shipmentQnty.clear();
					shipmentQnty.sendKeys(sQnty);
					 Thread.sleep(8000);//input[contains(@id,'shpQtyInput::content')]
				/*WebElement shipmentQnty= driver.findElement(By.xpath("//*[text()= '"+item+"']/../../../../../../..//input[contains(@id, 'inputText1::content')]"));
				shipmentQnty.clear();
				shipmentQnty.sendKeys(sQnty);
				 Thread.sleep(8000);
				 WebElement trak =driver.findElement(By.xpath("//span[text()='"+item+"']/../../../../../../../../..//input[contains(@id, 'inputText13::content')]"));
				 trak.click();
				 trak.clear();
				 trak.sendKeys(tNumb);*/
					 WebElement trak =driver.findElement(By.xpath("//input[contains(@id,'it5::content')]"));
					 trak.click();
					 trak.clear();
					 trak.sendKeys(tNumb);
				 Thread.sleep(5000);
				 driver.findElement(By.xpath("//*[contains(@id, 'ap1:WSHsv2::popEl')]")).click();
				 driver.findElement(By.xpath("//*[contains(text(), 'ave and Close')]")).click();
				 Thread.sleep(5000);
				 rowNumber++;
			 }
			
			 Thread.sleep(5000);
			 driver.findElement(By.xpath("//*[contains(@id, 'pt1:scap1:WSHsv2')]/table/tbody/tr/td[1]/a/span")).click();
			 Thread.sleep(5000);
			 driver.findElement(By.xpath("//*[contains(@id, 'pt1:scap1:shipconfirm')]/a/span")).click();
			 Thread.sleep(7000);
			 WebDriverWait wait11 = new WebDriverWait(driver, 500);
			 wait11.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id, 'pt1:scap1:panelFormLayout3')]/table/tbody/tr/td/table/tbody/tr[2]/td[2]")));
			 String shipmentAlert = driver.findElement(By.xpath("//*[contains(@id, 'pt1:scap1:panelFormLayout3')]/table/tbody/tr/td/table/tbody/tr[2]/td[2]")).getText();
	         String shipmentNum = shipmentAlert.replace("The shipment ", "");
	         String shipment = shipmentNum.replace(" was confirmed.", "").trim();
	         System.out.println("shipment="+shipment);
	         sheet.getRow(rowNumber-1).createCell(11).setCellValue(shipment);
	 		 fos = new FileOutputStream(srcFile);
	 		 wb.write(fos);
	 		 driver.findElement(By.xpath("//*[contains(@id, 'pt1:scap1:sccb31')]")).click();
		       Thread.sleep(5000);
		       WebDriverWait wait111 = new WebDriverWait(driver, 500);
		       wait111.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(@id, 'pt1:scap1:ctb1::popEl')]")));
		       driver.findElement(By.xpath("//a[contains(@id, 'pt1:scap1:ctb1::popEl')]")).click();
		       driver.findElement(By.xpath("//*[text()='Close']")).click();
		       Thread.sleep(5000);
		       driver.findElement(By.xpath("//button[contains(@id, 'pt1:scap1:cb6')]")).click();
		       Thread.sleep(5000);
		       sheet.getRow(rowNumber-1).createCell(12).setCellValue("Closed");
		       shipmentLink = shipmentLink+1;
		       CellStyle style = wb.createCellStyle();
				Font font = wb.createFont();
			    XSSFCell cell2 = sheet.getRow(rowNumber-1).createCell(13);
			    cell2.setCellValue("PASS");
			    font.setColor(IndexedColors.GREEN.getIndex());
				font.setBold(true);
				style = wb.createCellStyle();
				style.setFont(font);
				cell2.setCellStyle(style);
			    fos = new FileOutputStream(srcFile);
			    wb.write(fos);
	 		 
	        }
			 
            else if(lineStatus.equalsIgnoreCase("Backordered"))
			  {
				  shipmentLink = shipmentLink+1;
				  System.out.println("Line status is in Backordered");
			      Thread.sleep(4000);
			  }
			 else {
				  shipmentLink = shipmentLink+1;
				  System.out.println("Line status is not in Staged or Backordered");
				  Thread.sleep(4000);
//				  CellStyle style = wb.createCellStyle();
//					Font font = wb.createFont();
//				    XSSFCell cell2 = sheet.getRow(rowNumber).createCell(13);
//				    cell2.setCellValue("Fail");
//				    font.setColor(IndexedColors.RED.getIndex());
//					font.setBold(true);
//					style = wb.createCellStyle();
//					style.setFont(font);
//					cell2.setCellStyle(style);
//				    XSSFCell cell = sheet.getRow(rowNumber).createCell(14);
//				    cell.setCellValue("Line status is in....."+lineStatus);
//				    fos = new FileOutputStream(srcFile);
//				    wb.write(fos);
				   
			 }
			 }}
			 driver.findElement(By.xpath("//div[contains(@id, 'pt1:AP1:SPc')]")).click();
		}
				 else
				  {
				  rowNumber++;
				  }}
		 
			  }
		 else
		 {
			 System.out.println("File is already processed");
		 }
	}	
		catch(Exception e)
		{
			e.printStackTrace();
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
    
    public static boolean isRowEmpty(Row row) {
	    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	        Cell cell = row.getCell(c);
	        if (cell != null && cell.getCellType() != CellType.BLANK)
	            return false;
	    }
	    return true;
	  }

}
