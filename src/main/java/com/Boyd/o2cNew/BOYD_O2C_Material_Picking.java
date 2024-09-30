package com.Boyd.o2cNew;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import org.testng.Assert;
import io.github.bonigarcia.wdm.WebDriverManager;

@SuppressWarnings("unused")
public class BOYD_O2C_Material_Picking {
	
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
    WebDriver driver;
    File srcFile;
    XSSFSheet sheet;
    int i = 0;
    int rowNumber;
    @Test
	public void materialPick() throws Exception
	{
		try {
		  WebDriverManager.chromedriver().setup();
		  driver = new ChromeDriver();
		  driver.manage().deleteAllCookies();
		  driver.manage().timeouts().implicitlyWait(90, TimeUnit.SECONDS);
		  driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
		  driver.get("https://elme-dev1.fa.us8.oraclecloud.com/");
	//	  driver.get("https://elme-dev2.fa.us8.oraclecloud.com");
//		  driver.get("https://elme-test.login.us8.oraclecloud.com/");
		  driver.manage().window().maximize();
		  srcFile = new File(System.getProperty("user.dir")+"\\Excel\\BOYD_O2C_MaterialPicking.xlsx");
		  fis = new FileInputStream(srcFile);
		  wb = new XSSFWorkbook(fis);
		  sheet = wb.getSheet("CreatePickWave");
		  driver.findElement(By.id("userid")).sendKeys("forsys.user");
		  driver.findElement(By.id("password")).sendKeys("forsys2023");
   	//	  driver.findElement(By.id("password")).sendKeys("forsys4@4!");
		  driver.findElement(By.id("btnActive")).click();
		  Thread.sleep(5000);
		  WebDriverWait wait1 = new WebDriverWait(driver, 500);
		  wait1.until(ExpectedConditions.elementToBeClickable(By.id("pt1:_UIShome")));
		  driver.findElement(By.id("pt1:_UIShome")).click();
		  Thread.sleep(12000);
		  WebDriverWait wait = new WebDriverWait(driver, 500);
		  wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Supply Chain Execution")));
		  driver.findElement(By.linkText("Supply Chain Execution")).click();
		  driver.findElement(By.linkText("Inventory Management")).click();
		  int rowNum=sheet.getPhysicalNumberOfRows();
		  System.out.println("rowNum="+rowNum);
		  System.out.println("colNum="+sheet.getRow(1).getLastCellNum());
		  Row row = sheet.getRow(1);
		  Cell c = row.getCell(4);
		  System.out.println("result=="+c);
		 if(c==null||c.getStringCellValue().contentEquals(""))
			 {
		  
		  //**Create Pick Wave**//
		  for(rowNumber=1; rowNumber<rowNum; rowNumber++)
		  {
			  int orderNumber = (int) sheet.getRow(rowNumber).getCell(0).getNumericCellValue();
			  String order = String.valueOf(orderNumber).trim();
			  String warehouse = sheet.getRow(rowNumber).getCell(1).getStringCellValue().trim();
			  String relRule = sheet.getRow(rowNumber).getCell(2).getStringCellValue().trim();
		          driver.findElement(By.xpath("//img[contains(@id, 'itemNode_InvTasksList::icon')]")).click();
		          Thread.sleep(7000);
		          driver.findElement(By.xpath("//select[contains(@id, 'FOTRaT:0:soc1::content')]")).sendKeys("Shipments");
		          Thread.sleep(3000);
		          driver.findElement(By.linkText("Create Pick Wave")).click();
		          Thread.sleep(6000);
				  driver.findElement(By.xpath("//input[contains(@id, 'ap1:OrganizationCodeSL::content')]")).sendKeys(warehouse);
				  Thread.sleep(4000);
				  WebElement inputField=driver.findElement(By.xpath("//input[contains(@id, 'ap1:SalesOrderNumberSL::content')]"));
				  inputField.sendKeys(order);
				  Thread.sleep(9000);
				  driver.findElement(By.xpath("//input[@id='pt1:_FOr1:1:_FOSritemNode_warehouse_operations_InventoryManagement1:0:MAnt2:1:pt2:ap1:ToScheduledShipDateSL::content']")).clear();
				  driver.findElement(By.linkText("Show More")).click();
				  Thread.sleep(3000);
				  JavascriptExecutor executor = (JavascriptExecutor)driver;
				  executor.executeScript("window.scrollBy(0, 1000)", "");
				  Thread.sleep(3000);
				  driver.findElement(By.xpath("//input[@id='pt1:_FOr1:1:_FOSritemNode_warehouse_operations_InventoryManagement1:0:MAnt2:1:pt2:ap1:id2::content']")).clear();
				  Thread.sleep(1000);
				  executor.executeScript("window.scrollBy(0, -1000)", "");
				  Thread.sleep(3000);
				  driver.findElement(By.xpath("//button[@id='pt1:_FOr1:1:_FOSritemNode_warehouse_operations_InventoryManagement1:0:MAnt2:1:pt2:ap1:cb2']")).click();
				  Thread.sleep(8000);
				  String confirmationText = driver.findElement(By.xpath("//*[contains(@id, 'ap1:panelFormLayout12')]/table/tbody/tr/td/table/tbody/tr[2]/td[2]")).getText().trim();
				  Thread.sleep(5000);
				  driver.findElement(By.xpath("//button[contains(@id, 'ap1:cb1')]")).click();
				  System.out.println("confirmationText="+confirmationText);
				  String str = confirmationText.replace("Pick wave ", "").trim();
				  String pickWaveNumber = str.substring(0, 7).trim();
				  System.out.println("pickWaveNumber="+pickWaveNumber);
				  XSSFCell cell = sheet.getRow(rowNumber).createCell(3);
				  cell.setCellValue(pickWaveNumber);
				  fos = new FileOutputStream(srcFile);
				  wb.write(fos);
				  
					    CellStyle style = wb.createCellStyle();
						Font font = wb.createFont();
					    XSSFCell cell2 = sheet.getRow(rowNumber).createCell(4);
					    cell2.setCellValue("PASS");
					    font.setColor(IndexedColors.GREEN.getIndex());
						font.setBold(true);
						style = wb.createCellStyle();
						style.setFont(font);
						cell2.setCellStyle(style);
					    fos = new FileOutputStream(srcFile);
					    wb.write(fos);
					    try
					    {
					    	driver.findElement(By.xpath("//button[contains(@id, 'MAyes')]")).click();
					    }
					  catch(Exception e)
					    {
						  
					    }
				  
			   
		  }}
//}
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
			XSSFCell cell2 = sheet.getRow(rowNumber).createCell(4);
			cell2.setCellValue("Fail");
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style = wb.createCellStyle();
			style.setFont(font);
			cell2.setCellStyle(style);
			sheet.getRow(rowNumber).createCell(5).setCellValue("Exception occured");
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
		
//		driver.quit();
	}

}
