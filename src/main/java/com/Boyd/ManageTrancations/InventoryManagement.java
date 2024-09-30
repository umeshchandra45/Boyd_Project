package com.Boyd.ManageTrancations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
//import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class InventoryManagement extends Base_Class {
	public static WebDriverWait wait;
	public static int timeout = 60;
	public String transactionType;
	public String transactionNumber;
	public String interCompanyNumber;
	@Test
	public void Home_Page() throws InterruptedException, IOException
	{
		File f=new File(System.getProperty("user.dir")+"\\Excel\\UserAddOn.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Intercompany_Invoices");
		//sheet.getRow(0).createCell(3).setCellValue("Result");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		for(int i=1;i<totalrows;i++)
		{
			System.out.println("Count of i value :" +i);
			transactionType=sheet.getRow(i).getCell(1).getStringCellValue();
			DataFormatter df = new DataFormatter();
			transactionNumber = df.formatCellValue(sheet.getRow(i).getCell(0));
			interCompanyNumber=df.formatCellValue(sheet.getRow(i).getCell(2));
		WebElement tasks = browser.findElement(By.xpath("//img[contains(@title,'Tasks')]"));
		waitUntilElementClickable("tasks", tasks, browser, timeout);
		WebElement manageTransactions = browser.findElement(By.xpath("//a[text()='Manage Transactions']"));
		waitUntilElementClickable("manageTransactions", manageTransactions, browser, timeout);
		WebElement TransactionType = browser.findElement(By.xpath("//*[contains(@id,'pt1:MTF1:0:ap1:q1:value30::content')]"));
		waitUntilElementClickable("TransactionType", TransactionType, browser, timeout);
		TransactionType.sendKeys(transactionType);
		WebElement TransactionNumber = browser.findElement(By.xpath("//*[contains(@id,'pt1:MTF1:0:ap1:q1:value40::content')]"));
		waitUntilElementClickable("TransactionNumber", TransactionNumber, browser, timeout);
		TransactionNumber.sendKeys(transactionNumber);
		WebElement search = browser.findElement(By.xpath("//*[contains(@id,'pt1:MTF1:0:ap1:q1::search')]"));
		waitUntilElementClickable("search", search, browser, timeout);
		WebElement tNnumber = browser.findElement(By.xpath("//a[text()='"+transactionNumber+"']"));
		waitUntilElementClickable("tNnumber", tNnumber, browser, timeout);
		WebElement visibleText=browser.findElement(By.xpath("//*[contains(@title,'General Information')]"));
		visibleText.isDisplayed();
		WebElement action = browser.findElement(By.linkText("Actions"));
		waitUntilElementClickable("action", action, browser, timeout);
		/*String errormsg1 = browser.findElement(By.xpath("//td[contains(@class,'xp8')]")).getText().trim();
	    String str1[] = errormsg1.split("\\.");
        System.out.println(str1[0]);*/
        try {
		WebElement viewAccounting = browser.findElement(By.xpath("//td[text()='View Accounting']"));
		waitUntilElementClickable("viewAccounting", viewAccounting, browser, timeout);
		WebElement overrideAccount = browser.findElement(By.xpath("//*[contains(@id,'pt1:MTF1:1:pt1:Trans1:0:ap110:ViewA1:0:AT1:_ATp:cb34')]"));
	 	waitUntilElementClickable("overrideAccount", overrideAccount, browser, 5);  	
		WebElement newAccountSearch = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:ViewA1:0:r1:1:kf1KBIMG::icon')]"));
		waitUntilElementClickable("newAccountSearch", newAccountSearch, browser, timeout);
		Thread.sleep(9000);
		browser.findElement(By.xpath("//div[text()='New Account']")).click();
		Thread.sleep(3000);
		WebElement interCompapny = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:ViewA1:0:r1:1:kf1SPOP_query:value30::lovIconId')]"));
		waitUntilElementClickable("interCompapny", interCompapny, browser, timeout);
		Thread.sleep(6000);
		WebElement search2 = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("search2", search2, browser, timeout);
		WebElement value2 = browser.findElement(By.xpath("//*[contains(@id,'value30::_afrLovInternalQueryId:value00::content')]"));
		waitUntilElementClickable("value2", value2, browser, timeout);
		value2.clear();
		WaituntilElementwritable("value2", value2, browser, interCompanyNumber);
		browser.findElement(By.xpath("//*[contains(@id,'afrLovInternalQueryId::search')]")).click();
		WebElement tableValue = browser.findElement(By.xpath("//*[contains(@id,'afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("tableValue", tableValue, browser, timeout);
		WebElement ok = browser.findElement(By.xpath("//*[contains(@id,'Trans1:0:ap110:ViewA1:0:r1:1:kf1SPOP_query:value30::lovDialogId::ok')]"));
		waitUntilElementClickable("ok", ok, browser, timeout);
		WebElement ok1 = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:ViewA1:0:r1:1:kf1SEl')]"));
		waitUntilElementClickable("ok1", ok1, browser, timeout);
		WebElement reason = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:ViewA1:0:r1:1:it1::content')]"));
		waitUntilElementClickable("reason", reason, browser, timeout);
		reason.sendKeys("test");
		try {
		WebElement submit = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:ViewA1:0:commandButton1')]"));
		waitUntilElementClickable("submit", submit, browser, 10);
		Thread.sleep(5000);
		WebElement done = browser.findElement(By.xpath("//*[contains(@id,'MAnt2:1:pt1:MTF1:1:pt1:Trans1:0:ap110:d12::ok')]"));
		waitUntilElementClickable("done", done, browser, timeout);
		WebElement save = browser.findElement(By.xpath("//*[contains(@id,'MAnt2:1:pt1:MTF1:1:pt1:Trans1:0:ap110:saveMenu::popEl')]"));
		waitUntilElementClickable("save", save, browser, timeout);
		WebElement saveAndClose = browser.findElement(By.xpath("//td[text()='ave and Close']"));
		waitUntilElementClickable("saveAndClose", saveAndClose, browser, timeout);
		WebElement ok3 = browser.findElement(By.xpath("//*[contains(@id,'d1::msgDlg::cancel')]"));
		waitUntilElementClickable("ok3", ok3, browser, timeout);
		WebElement done2 = browser.findElement(By.xpath("//*[contains(@id,'0:MAnt2:1:pt1:MTF1:0:ap1:cb1')]"));
		waitUntilElementClickable("done2", done2, browser, timeout);
		sheet.getRow(i).createCell(3).setCellValue("Updated");
		Updatefile(f, wb);
		System.out.println("intercompany updated for "+transactionNumber.toString());
		}
		catch(Exception e)
		{
		    String errormsg = browser.findElement(By.xpath("//td[contains(@class,'x1mz')]")).getText().trim();
		    String str[] = errormsg.split("\\.");
	        System.out.println(str[0]);
	       // sheet.getRow(0).createCell(4).setCellValue("comments");
			sheet.getRow(i).createCell(4).setCellValue(str[0]);
			sheet.getRow(i).createCell(3).setCellValue("fail");
			Updatefile(f, wb);
			WebElement error = browser.findElement(By.xpath("//*[contains(@id,'d1::msgDlg::cancel')]"));
			waitUntilElementClickable("error", error, browser, timeout);
			WebElement cancel = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:ViewA1:0:cb3')]"));
			waitUntilElementClickable("cancel", cancel, browser, timeout);
			WebElement done3 = browser.findElement(By.xpath("//*[contains(@id,'pt1:Trans1:0:ap110:d12::ok')]"));
			waitUntilElementClickable("done3", done3, browser, timeout);
			WebElement cancel2 = browser.findElement(By.xpath("//*[contains(@id,'ap110:commandToolbarButton2')]"));
			waitUntilElementClickable("cancel2", cancel2, browser, timeout);
			WebElement yes = browser.findElement(By.xpath("//*[contains(@id,'dialogCancel::yes')]"));
			waitUntilElementClickable("yes", yes, browser, timeout);
			WebElement done4 = browser.findElement(By.xpath("//*[contains(@id,'MAnt2:1:pt1:MTF1:0:ap1:cb1')]"));
			waitUntilElementClickable("done4", done4, browser, timeout);
		}
	}
        
		catch(Exception e)
		{
			System.out.println("exception for sumbmit");
		//	sheet.getRow(0).createCell(4).setCellValue("comments");
			sheet.getRow(i).createCell(4).setCellValue("Required paymentTerms");
			sheet.getRow(i).createCell(3).setCellValue("failed");
			Updatefile(f, wb);
			WebElement cancel6 = browser.findElement(By.xpath("//*[contains(@id,'ap110:commandToolbarButton2')]"));
			waitUntilElementClickable("cancel6", cancel6, browser, timeout);
			WebElement yes1 = browser.findElement(By.xpath("//*[contains(@id,'dialogCancel::yes')]"));
			waitUntilElementClickable("yes1", yes1, browser, timeout);
			WebElement done4 = browser.findElement(By.xpath("//*[contains(@id,'MAnt2:1:pt1:MTF1:0:ap1:cb1')]"));
			waitUntilElementClickable("done4", done4, browser, timeout);

		}
	

	try
	{
		wb.close();
	}
	catch(Exception e)
	{
		
	}}}
	
	
	
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


}
