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
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;
public class BOYD_O2C_Confirm_Picking {
	
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
    WebDriver driver;
    File srcFile;
    XSSFSheet sheet;
    int i = 0;
    int rowNumber;
	int count = 0;
	int pickslips = 0;
    @SuppressWarnings("unused")
	@Test
	public void materialPick() throws Exception
	{
		try {
		  WebDriverManager.chromedriver().setup();
		  driver = new ChromeDriver();
		  driver.manage().deleteAllCookies();
		  driver.manage().timeouts().implicitlyWait(80, TimeUnit.SECONDS);
		  driver.manage().timeouts().pageLoadTimeout(80, TimeUnit.SECONDS);
		  driver.get("https://elme-dev1.fa.us8.oraclecloud.com/");
//		  driver.get("https://elme-dev2.fa.us8.oraclecloud.com");
//		  driver.get("https://elme-test.login.us8.oraclecloud.com/");
		  driver.manage().window().maximize();
		  srcFile = new File(System.getProperty("user.dir")+"\\Excel\\BOYD_O2C_MaterialPicking.xlsx");
		  fis = new FileInputStream(srcFile);
		  wb = new XSSFWorkbook(fis);
		  sheet = wb.getSheet("ConfirmPicking");
		  driver.findElement(By.id("userid")).sendKeys("forsys.user");
		  driver.findElement(By.id("password")).sendKeys("forsys2023");
   		 // driver.findElement(By.id("password")).sendKeys("forsys4@4!");
		  driver.findElement(By.id("btnActive")).click();
		  JavascriptExecutor js = (JavascriptExecutor) driver;
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
		  Row row = sheet.getRow(1);
		  Cell c = row.getCell(10);
		  System.out.println("result=="+c);
		 if(c==null||c.getStringCellValue().contentEquals(""))
			 {
		  
		  //**Confirm Pick Slips **//
		  
		  for(rowNumber=1; rowNumber<rowNum;)
		  {
			  if(sheet.getRow(rowNumber) == null || isRowEmpty(sheet.getRow(rowNumber))) {
					 
				        Thread.sleep(3000);
					    rowNumber++;
					    continue;
			  }
			  
			  String order = sheet.getRow(rowNumber).getCell(0).getStringCellValue().trim();
			  String pickWave = sheet.getRow(rowNumber).getCell(1).getStringCellValue().trim();
			String date = sheet.getRow(rowNumber).getCell(2).getStringCellValue().trim();
			  
			  if(!order.equalsIgnoreCase("NA"))
			  {
			  driver.findElement(By.xpath("//img[contains(@id, 'itemNode_InvTasksList::icon')]")).click();
	          Thread.sleep(3000);
	          driver.findElement(By.xpath("//select[contains(@id, 'FOTRaT:0:soc1::content')]")).sendKeys("Shipments");
	          Thread.sleep(2500);
	          driver.findElement(By.linkText("Confirm Pick Slips")).click();
	          Thread.sleep(5000);
		  driver.findElement(By.xpath("//input[contains(@id, 'ap1:q1:value20::content')]")).sendKeys(order);
		  driver.findElement(By.xpath("//input[contains(@id, 'ap1:q1:value40::content')]")).sendKeys(pickWave);
		  Thread.sleep(3000);
		  driver.findElement(By.xpath("//input[contains(@id, 'ap1:q1:value60::content')]")).click();
		  driver.findElement(By.xpath("//input[contains(@id, 'ap1:q1:value60::content')]")).clear();
		  driver.findElement(By.xpath("//button[contains(@id, 'ap1:q1::search')]")).click();
		  Thread.sleep(8000);
		  js.executeScript("window.scrollBy(0, 2000)", "");
		  Thread.sleep(3000);
		  List<WebElement> table =driver.findElements(By.xpath("//*[contains(@id, 'ap1:AT1:_ATp:t2::db')]/table/tbody/tr"));
			int size = table.size();
			System.out.println("Size of the table =="+size);
			if(size==0)
			{
				System.out.println("No data found with the given order");
				driver.findElement(By.xpath("//*[contains(@id, 'pt1:ap1:SPb')]")).click();
				 CellStyle style = wb.createCellStyle();
					Font font = wb.createFont();
				    XSSFCell cell2 = sheet.getRow(rowNumber).createCell(10);
				    cell2.setCellValue("Fail");
				    font.setColor(IndexedColors.RED.getIndex());
					font.setBold(true);
					style = wb.createCellStyle();
					style.setFont(font);
					cell2.setCellStyle(style);
					sheet.getRow(rowNumber).createCell(11).setCellValue("No data found with the given order");
				    fos = new FileOutputStream(srcFile);
				    wb.write(fos);
				    rowNumber++;
				  
			}
			else
			{
				
				int k=0;
			while(k<size)
			{
				driver.findElement(By.xpath("//*[text()= '"+order+"']/../../../../../../..//a[contains(@id, 'cl1')]")).click();
				Thread.sleep(5000);
//		  String openPicks = driver.findElement(By.xpath("//*[@id=\'pt1:_FOr1:1:_FOSritemNode_warehouse_operations_InventoryManagement1:0:MAnt2:2:pt1:ap1:AT1:_ATp:t2::db\']/table/tbody/tr/td[3]/div/table/tbody/tr/td[8]/span")).getText();
//		  System.out.println("openPicks="+openPicks);
//		  int picks = Integer.parseInt(openPicks);
//		  driver.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_warehouse_operations_InventoryManagement1:0:MAnt2:2:pt1:ap1:AT1:_ATp:t2:"+pickslips+":cl1")).click();
//		  count = count+picks;
		  
		 List<WebElement> table1 =driver.findElements(By.xpath("//*[contains(@id, 'pt1:apppanel:AT1:_ATp:table1::db')]/table/tbody/tr"));
		 int size1 = table1.size();
		 System.out.println("Size of the table =="+size1);
		 int m=0;
		 while(m<size1)
		 {
			 
			 String itemNum = sheet.getRow(rowNumber).getCell(3).getStringCellValue().trim();
			  int lineNumber = (int) sheet.getRow(rowNumber).getCell(4).getNumericCellValue();
			String line = String.valueOf(lineNumber).trim();
			  int reqQnty = (int) sheet.getRow(rowNumber).getCell(5).getNumericCellValue();
			  String rQnty = String.valueOf(reqQnty).trim();
			  int pickQnty = (int) sheet.getRow(rowNumber).getCell(6).getNumericCellValue();
			  String pQnty = String.valueOf(pickQnty).trim();
			  String subInventory = sheet.getRow(rowNumber).getCell(7).getStringCellValue().trim();
			  String locator = sheet.getRow(rowNumber).getCell(8).getStringCellValue().trim();
			  String lot = sheet.getRow(rowNumber).getCell(9).getStringCellValue().trim();
			  
			  driver.findElement(By.xpath("//*[text()= '"+itemNum+"']")).click();
			  Thread.sleep(4000);
			  driver.findElement(By.xpath("//*[text()= '"+itemNum+"']/../../../../../../../../..//label[contains(@id, 'sbc1::Label0')]")).click();
			  Thread.sleep(4000);
			  WebElement appQnty = driver.findElement(By.xpath("//*[text()= '"+itemNum+"']/../../../..//input[contains(@id, 'pickedqtyid::content')]"));
			  appQnty.click();
			  Thread.sleep(3000);
			  appQnty.sendKeys(Keys.chord(Keys.SHIFT,Keys.END));
			  appQnty.sendKeys(Keys.BACK_SPACE);
			  appQnty.sendKeys(pQnty);
			  appQnty.sendKeys(Keys.TAB);
//			  appQnty.clear();
////			  Thread.sleep(3000);
//			  WebElement appQnty1 = driver.findElement(By.xpath("//*[text()= '"+itemNum+"']/../../../..//input[contains(@id, 'pickedqtyid::content')]"));
//			  appQnty1.sendKeys(pQnty);
////			  appQnty1.sendKeys(Keys.ENTER);
			  Thread.sleep(5000);
//			  if(reqQnty>pickQnty)
//			  {
//				  driver.findElement(By.xpath("//button[contains(@id, 'pt1:apppanel:AT1:yesbutton')]")).click();
//				  Thread.sleep(5000);
//				  driver.findElement(By.xpath("//*[text()= '"+itemNum+"']/../../../..//a[contains(@id, 'lotnumberid::lovIconId')]")).click();
//				  Thread.sleep(4000);
//				  driver.findElement(By.xpath("//*[contains(@id, 'dropdownPopup::dropDownContent::db')]/table/tbody/tr[2]/td[1]")).click();
//				  Thread.sleep(3000);
//			  }
		 Thread.sleep(5000);
		 CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
		    XSSFCell cell2 = sheet.getRow(rowNumber).createCell(10);
		    cell2.setCellValue("Pass");
		    font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style = wb.createCellStyle();
			style.setFont(font);
			cell2.setCellStyle(style);
		    fos = new FileOutputStream(srcFile);
		    wb.write(fos);
		    m++;
			rowNumber++;
		 }
		 driver.findElement(By.xpath("//a[contains(@id, 'pt1:apppanel:SPsb::popEl')]")).click();
		 Thread.sleep(3000);
		 driver.findElement(By.xpath("//*[contains(@id, 'pt1:apppanel:cmi2')]")).click();
		 Thread.sleep(5000);
		 driver.findElement(By.xpath("//*[contains(@id, 'pt1:ap1:SPb')]")).click();
		 k++;
		 }
			
			 
	}
			  
		  }
			  else
				  {
				  rowNumber++;
				  }
				  }
}
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
		    XSSFCell cell2 = sheet.getRow(rowNumber).createCell(10);
		    cell2.setCellValue("Fail");
		    font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style = wb.createCellStyle();
			style.setFont(font);
			cell2.setCellStyle(style);
			sheet.getRow(rowNumber).createCell(11).setCellValue("Exception occured");
		    fos = new FileOutputStream(srcFile);
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
    public static boolean isRowEmpty(Row row) {
	    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	        Cell cell = row.getCell(c);
	        if (cell != null && cell.getCellType() != CellType.BLANK)
	            return false;
	    }
	    return true;
	  }

}
