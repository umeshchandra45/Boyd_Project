package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FileUtils;
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
import org.openqa.selenium.OutputType;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class BOYD_P2P_Invoice_Generation {
	
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	File f;
	WebDriver driver;
	public static String bussUnit;
	public static String poNum;
	public static String supSite;
    public static String invoiceGrp;
    public static String invoiceNum;
    public static String payTerms;
    public static String amount;
    int i;
	
	@BeforeTest()
	 public void beforeTest() throws Exception {
	 WebDriverManager.chromedriver().setup();
	 ChromeOptions options = new ChromeOptions();
	 options.setPageLoadStrategy(PageLoadStrategy.NONE);
	 driver = new ChromeDriver();
	 driver.manage().window().maximize();
	 driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
	 driver.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
	driver.get("https://elme-dev1.fa.us8.oraclecloud.com");
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
		//	driver.findElement(By.id("password")).sendKeys("forsys4@4!");
	 driver.findElement(By.id("btnActive")).click();
	 Thread.sleep(7000);
	 driver.findElement(By.id("pt1:_UIShome")).click();
	 Thread.sleep(10000);
	 driver.findElement(By.linkText("Payables")).click();
	 driver.findElement(By.linkText("Invoices")).click();
}
	
	@Test()
	 public void invoiceMatching() throws IOException {
		try
		{
		
			String remoteExecution = System.getProperty("remoteExecution");
			System.out.println("remoteFlag  is :" +remoteExecution);
			Boolean exectionFlag = Boolean.parseBoolean(remoteExecution);
			// File f = null;
			if(exectionFlag) {
//				System.out.println("remote execution");
//				f = new File(System.getProperty("user.dir") + "\\ExternalFiles\\BOYD_P2P_InvoiceGeneration.xlsx");
			}			 
			else {
				 f = new File(System.getProperty("user.dir") + "\\Excel\\BOYD_P2P_InvoiceGeneration.xlsx");
//				 System.out.println("local execution");
			}
			
			 fis = new FileInputStream(f);
			 wb = new XSSFWorkbook(fis);
			 sheet = wb.getSheet("Invoice_Generation");
			 int totalRows = sheet.getPhysicalNumberOfRows();
			 JavascriptExecutor js = (JavascriptExecutor) driver;
			 System.out.println("Total number of Excel rows are :" +totalRows);
			 Row row = sheet.getRow(1);
			  Cell c = row.getCell(8);
			  System.out.println("result=="+c);
			 if(c==null||c.getStringCellValue().contentEquals(""))
			 {
			 for(i=1; i<=totalRows; i++) 
			 {
				 if(sheet.getRow(i) == null || isRowEmpty(sheet.getRow(i))) {
					 
					 continue;
				 }
			 bussUnit = sheet.getRow(i).getCell(0).getStringCellValue().trim();
			poNum = sheet.getRow(i).getCell(1).getStringCellValue().trim();
		   supSite  = sheet.getRow(i).getCell(2).getStringCellValue().trim();
		   invoiceGrp = sheet.getRow(i).getCell(3).getStringCellValue().trim();
		   invoiceNum = sheet.getRow(i).getCell(4).getStringCellValue().trim();
			 payTerms = sheet.getRow(i).getCell(5).getStringCellValue().trim();
			 amount = sheet.getRow(i).getCell(6).getStringCellValue().trim();
			 
			driver.findElement(By.xpath("//img[contains(@id, 'FndTasksList')]")).click();
			driver.findElement(By.linkText("Create Invoice")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//a[contains(@id, 'ic1::lovIconId')]")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//input[contains(@id, 'ic1::_afrLovInternalQueryId:value00::content')]")).sendKeys(poNum);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//button[contains(@id, 'ic1::_afrLovInternalQueryId::search')]")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//*[contains(@id, 'ic1_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
			driver.findElement(By.xpath("//button[contains(@id, 'ic1::lovDialogId::ok')]")).click();
			Thread.sleep(5000);
			if(!supSite.equalsIgnoreCase("NA") && !supSite.isEmpty())
			  {
			WebElement site = driver.findElement(By.xpath("//input[contains(@id, 'ic4::content')]"));
			site.clear();
			site.sendKeys(supSite);
			site.sendKeys(Keys.TAB);
			Thread.sleep(5000);
			  }
			  Thread.sleep(6000);                                    
			driver.findElement(By.xpath("//input[contains(@id, 'i1::content')]")).click();
			driver.findElement(By.xpath("//input[contains(@id, 'i1::content')]")).sendKeys(invoiceGrp);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[contains(@id, 'i2::content')]")).sendKeys(invoiceNum);
			Thread.sleep(3000);
			WebElement ele = driver.findElement(By.xpath("//input[contains(@id, 'i3::content')]"));
			ele.sendKeys(amount);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[contains(@id, 'so3::content')]")).click();
			driver.findElement(By.xpath("//input[contains(@id, 'so3::content')]")).clear();
			driver.findElement(By.xpath("//input[contains(@id, 'so3::content')]")).sendKeys(payTerms);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//img[contains(@id, 'ap1:cg1::icon')]")).click();
			Thread.sleep(6000);
			try
			{
				WebElement error = driver.findElement(By.id("d1::msgDlg::_ttxt"));
				String err = error.getText();
				System.out.println("error=="+err);
				if(err.contains("Error"))
				{
					driver.findElement(By.id("d1::msgDlg::cancel")).click();
					CellStyle style = wb.createCellStyle();  
				     XSSFCell cell1 = sheet.getRow(i).createCell(8);
				     cell1.setCellValue("Fail");
					 Font font = wb.createFont();
					 font.setColor(IndexedColors.RED.getIndex());
					 font.setBold(true);
					 style = wb.createCellStyle();
					 style.setFont(font);
					 cell1.setCellStyle(style);
					 sheet.getRow(i).createCell(9).setCellValue("This invoice number already exists. Enter a unique invoice number");
					 fos = new FileOutputStream(f);
					 wb.write(fos);
					 driver.findElement(By.xpath("//span[text()='ancel']")).click();
					 break;
					
				}
				
			}
			catch(Exception ex)
			{
				
			}
//			driver.findElement(By.xpath("//label[contains(@id, 'at1:_ATp:ta1:0:sb1::Label0')]")).click();
			driver.findElement(By.xpath("//label[contains(@id, 'sb3::Label0')]")).click();
			try {
				Thread.sleep(5000);
//				driver.findElement(By.xpath("//button[contains(@id, 'at1:_ATp:ta1:0:cb3')]")).click();
				driver.findElement(By.xpath("//button[contains(@id, 'ap1:r11:1:cb31')]")).click();
			}
			catch(Exception exe)
			{
				
			}
			Thread.sleep(8000);
//			WebElement amt = driver.findElement(By.xpath("//input[contains(@id, 'ATp:ta1:0:i3::content')]"));
			WebElement amt = driver.findElement(By.xpath("//span[contains(@id, 'at1:_ATp:ta1:o268')]"));
			String amt1 = amt.getText();
			System.out.println("actual amount=="+amt1);
			sheet.getRow(i).createCell(7).setCellValue(amt1);
			driver.findElement(By.xpath("//button[contains(@id, 'pm1:r1:0:ap1:cb2')]")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//button[contains(@id, 'pm1:r1:0:ap1:cb17')]")).click();
			Thread.sleep(6000);
			WebDriverWait wait = new WebDriverWait(driver,350);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[contains(@id, 'i3::content')]")));
			driver.findElement(By.xpath("//input[contains(@id, 'i3::content')]")).click();
			driver.findElement(By.xpath("//input[contains(@id, 'i3::content')]")).clear();
			driver.findElement(By.xpath("//input[contains(@id, 'i3::content')]")).sendKeys(amt1);
			Thread.sleep(4000);
			driver.findElement(By.xpath("//span[text()='Save']")).click();
			Thread.sleep(12000);
			driver.findElement(By.linkText("Invoice Actions")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//td[text()='Validate']")).click();
			Thread.sleep(20000);
			driver.findElement(By.xpath("//button[contains(@id, 'msgDlg::cancel')]")).click();
			String text = driver.findElement(By.xpath("//a[contains(@id, 'ap1:cl3')]")).getText();
			System.out.println("text===="+text);
			if(text.equalsIgnoreCase("Validated"))
			{

				Thread.sleep(3000);
				driver.findElement(By.linkText("Invoice Actions")).click();
				Thread.sleep(4000);
				driver.findElement(By.xpath("//td[text()='Post to Ledger']")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//button[contains(@id, 'ap1:cb42')]")).click();
				Thread.sleep(5000);
				captureScreenShot(driver);
				List<WebElement> ele11 = driver.findElements(By.xpath("//*[contains(@id, 'ap1:r7:1:AT1:_ATp:t1::db')]/table/tbody/tr"));
				int eleSize = ele11.size();
				System.out.println("Size of table =="+eleSize);
				for(int k=0; k<eleSize; k++)
				{
					WebElement rows = driver.findElement(By.xpath("//*[contains(@id, 'ap1:r7:1:AT1:_ATp:t1::db')]/table/tbody/tr["+(k+1)+"]"));
					String rowData = rows.getText();
					System.out.println("rowData=="+rowData);
					WebElement debit = driver.findElement(By.xpath("//*[contains(@id, 'ap1:r7:1:AT1:_ATp:t1::db')]/table/tbody/tr["+(k+1)+"]/td[7]"));
					String debt = debit.getText();
					System.out.println("debit=="+debt);
					if(debt.equals(" "))
					{
						Thread.sleep(2000);
						sheet.getRow(0).createCell(10+k).setCellValue("Credit");
						CellStyle style = wb.createCellStyle();  
					     XSSFCell cell1 = sheet.getRow(i).createCell(10+k);
					     cell1.setCellValue(rowData);
						 Font font = wb.createFont();
						 font.setColor(IndexedColors.ORANGE.getIndex());
						 font.setBold(true);
						 style = wb.createCellStyle();
						 style.setFont(font);
						 cell1.setCellStyle(style);
						 fos = new FileOutputStream(f);
						 wb.write(fos);
					}
					else
					{
						Thread.sleep(2000);
						sheet.getRow(0).createCell(10+k).setCellValue("Debit");
						CellStyle style = wb.createCellStyle();  
					     XSSFCell cell1 = sheet.getRow(i).createCell(10+k);
					     cell1.setCellValue(rowData);
						 Font font = wb.createFont();
						 font.setColor(IndexedColors.BLUE.getIndex());
						 font.setBold(true);
						 style = wb.createCellStyle();
						 style.setFont(font);
						 cell1.setCellStyle(style);
						 fos = new FileOutputStream(f);
						 wb.write(fos);
					}
					
				}
				Thread.sleep(5000);
				driver.findElement(By.xpath("//button[contains(@id, 'ap1:cb99')]")).click();
				Thread.sleep(7000);
				driver.findElement(By.xpath("//span[text()='ave and Close']")).click();
				Thread.sleep(5000);
				 CellStyle style = wb.createCellStyle();  
			     XSSFCell cell1 = sheet.getRow(i).createCell(8);
			     cell1.setCellValue("Pass");
				 Font font = wb.createFont();
				 font.setColor(IndexedColors.GREEN.getIndex());
				 font.setBold(true);
				 style = wb.createCellStyle();
				 style.setFont(font);
				 cell1.setCellStyle(style);
				fos = new FileOutputStream(f);
				wb.write(fos);
				driver.findElement(By.xpath("//button[contains(@id, 'msgDlg::cancel')]")).click();
				
				//Approval Process//
				driver.findElement(By.xpath("//img[contains(@id, 'FndTasksList')]")).click();
				driver.findElement(By.linkText("Initiate Approval Workflow")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_BUSINESSUNIT::content')]")).click();
				driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_BUSINESSUNIT::content')]")).clear();
				driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_BUSINESSUNIT::content')]")).sendKeys(bussUnit);
				Thread.sleep(4000);
				driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_INVOICENUM::content')]")).click();
				driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_INVOICENUM::content')]")).clear();
				Thread.sleep(3000);
				driver.findElement(By.xpath("//input[contains(@id, 'paramDynForm_INVOICENUM::content')]")).sendKeys(invoiceNum);
				Thread.sleep(4000);
				driver.findElement(By.xpath("//div[contains(@id, 'requestBtns:submitButton')]")).click();
				Thread.sleep(7000);
				driver.findElement(By.xpath("//button[contains(@id, 'confirmSubmitDialog::ok')]")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//div[contains(@id, 'requestBtns:cancelButton')]")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//img[contains(@id, 'FndTasksList')]")).click();
				driver.findElement(By.linkText("Manage Invoices")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//input[contains(@id, 'value10::content')]")).sendKeys(invoiceNum);
				Thread.sleep(7000);
				driver.findElement(By.xpath("//button[contains(@id, 'ap1:q1::search')]")).click();
				Thread.sleep(5000);
//				WebElement appsts = driver.findElement(By.xpath("//*[contains(@id, 'at1:_ATp:ta1::db')]/table/tbody/tr/td[2]/div/table/tbody/tr/td[15]/span"));
//				js.executeScript("arguments[0].scrollIntoView();",appsts );
//				Thread.sleep(5000);
				driver.findElement(By.xpath("//div[contains(@id, 'ap1:ctb2')]")).click();
				Thread.sleep(4000);
			}
			else
			{
				driver.findElement(By.xpath("//span[text()='ancel']")).click();
				CellStyle style = wb.createCellStyle();  
			     XSSFCell cell1 = sheet.getRow(i).createCell(8);
			     cell1.setCellValue("Fail");
				 Font font = wb.createFont();
				 font.setColor(IndexedColors.RED.getIndex());
				 font.setBold(true);
				 style = wb.createCellStyle();
				 style.setFont(font);
				 cell1.setCellStyle(style);
				 fos = new FileOutputStream(f);
				 wb.write(fos);
			}
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
			     XSSFCell cell1 = sheet.getRow(i).createCell(8);
			     cell1.setCellValue("Fail");
				 Font font = wb.createFont();
				 font.setColor(IndexedColors.RED.getIndex());
				 font.setBold(true);
				 style = wb.createCellStyle();
				 style.setFont(font);
				 cell1.setCellStyle(style);
				 sheet.getRow(i).createCell(9).setCellValue("Exception Occured");
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
	public void captureScreenShot(WebDriver driver)
	{
		
		try
		{
			File src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(src, new File("D:\\Padma\\ERPTesting\\ScreenShots\\2Way_Matching"+System.currentTimeMillis()+".png"));
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
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
