package com.Boyd.ManageTrancations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Create_Customer_Site extends Base_Class {
	public static WebDriverWait wait;
	public static int timeout = 60;
	public String billtoAccountNumber;
	public String siteNumber;
	public String paymentTermsData;
	@Test
	public void Home_Page() throws InterruptedException, IOException
	{
		File f=new File(System.getProperty("user.dir")+"\\Excel\\CustomerSite_creation.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("sheet1");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		WebElement tasks = browser.findElement(By.xpath("//img[contains(@title,'Tasks')]"));
		waitUntilElementClickable("tasks", tasks, browser, timeout);
		WebElement manageCustomers = browser.findElement(By.xpath("//a[text()='Manage Customers']"));
		waitUntilElementClickable("manageCustomers", manageCustomers, browser, timeout);
		for(int i=1;i<totalrows;i++)
		{
		System.out.println("Count of i value :" +i);
		DataFormatter df = new DataFormatter();
		billtoAccountNumber = df.formatCellValue(sheet.getRow(i).getCell(0));
		System.out.println(billtoAccountNumber);
		siteNumber=sheet.getRow(i).getCell(1).getStringCellValue();
		System.out.println(siteNumber);
		paymentTermsData=sheet.getRow(i).getCell(2).getStringCellValue();
		System.out.println(paymentTermsData);
		Thread.sleep(2000);
		WebElement accountNumber = browser.findElement(By.xpath("//input[contains(@id,'AP1:q1:value120::content')]"));
		waitUntilElementClickable("accountNumber", accountNumber, browser, timeout);
		accountNumber.clear();
		accountNumber.sendKeys(billtoAccountNumber);
		WebElement search = browser.findElement(By.xpath("//*[contains(@id,':q1::search')]"));
		waitUntilElementClickable("search", search, browser, timeout);
		Thread.sleep(3000);
		WebElement siteInput = browser.findElement(By.xpath("//*[contains(@id,'ATp_afr_table1_afr_column66::content')]"));
		waitUntilElementClickable("siteInput", siteInput, browser, timeout);
		siteInput.sendKeys(siteNumber);
		siteInput.sendKeys(Keys.ENTER);
		Thread.sleep(2000);
		WebElement siteText = browser.findElement(By.xpath("//a[text()='" + siteNumber + "']"));
		waitUntilElementClickable("siteText", siteText, browser, timeout);
	
		WebElement profileHistory = browser.findElement(By.xpath("//div[@class='x1gb']//a[text()='Profile History']"));
		waitUntilElementClickable("profileHistory", profileHistory, browser, timeout);
		//Thread.sleep(2000);
		try {
		WebElement createSiteProfile = browser.findElement(By.xpath("//button[text()='Create Site Profile']"));
		waitUntilElementClickable("createSiteProfile", createSiteProfile, browser, timeout);
	/*	WebElement checkBox = browser.findElement(By.xpath("//input[contains(@id,'AP2:cuconsInvFlag::content')]"));
		waitUntilElementClickable("checkBox", checkBox, browser, timeout);
		boolean isSelected = checkBox.isSelected();         
		Thread.sleep(3000);
		if(isSelected==false) {
			waitUntilElementClickable("checkBox", checkBox, browser, timeout);
			Thread.sleep(4000);
			WebElement yes = browser.findElement(By.xpath("//*[contains(@id,'CManF:2:AP2:cb4')]"));
			waitUntilElementClickable("yes", yes, browser, timeout);
			Select billLevel =new Select(browser.findElement(By.xpath("//*[contains(@id,'AP2:soc10::content')]")));
			billLevel.selectByValue("Account");
			Select billType =new Select(browser.findElement(By.xpath("//*[contains(@id,'AP2:soc7::content')]")));
			billType.selectByValue("Detail");
			WebElement paymentTermsText = browser.findElement(By.xpath("//*[contains(@id,'AP2:bfbTermName::content')]"));
			waitUntilElementClickable("paymentTermsText", paymentTermsText, browser, timeout);
			paymentTermsText.clear();
			waitUntilElementClickable("paymentTermsText", paymentTermsText, browser, timeout);
			paymentTermsText.sendKeys(paymentTermsData);
			Thread.sleep(3000);
			}*/
		Thread.sleep(4000);
		WebElement saveAndClose = browser.findElement(By.xpath("//button[text()='ave and Close']"));
		waitUntilElementClickable("saveAndClose", saveAndClose, browser, timeout);
		Thread.sleep(3000);
		try {
			WebElement saveAndClose1 = browser.findElement(By.xpath("//button[text()='ave and Close']"));
			waitUntilElementClickable("saveAndClose1", saveAndClose1, browser, timeout);
			WebElement searchExpand = browser.findElement(By.xpath("//*[contains(@id,'AP1:q1::_afrDscl')]"));
			waitUntilElementClickable("searchExpand", searchExpand, browser, timeout);
			System.out.println("================created===================");
			sheet.getRow(i).createCell(3).setCellValue("created");
			Updatefile(f, wb);
		}
		catch(Exception e){
			browser.findElement(By.xpath("//button[contains(@id,'d1::msgDlg::cancel')]")).click();
			Thread.sleep(2000);
			browser.findElement(By.xpath("//button[contains(@id,'AP2:cb3')]")).click();
			Thread.sleep(2000);
			browser.findElement(By.xpath("//button[text()='ave and Close']")).click();
			Thread.sleep(4000);
			WebElement searchExpand = browser.findElement(By.xpath("//*[contains(@id,'AP1:q1::_afrDscl')]"));
			waitUntilElementClickable("searchExpand", searchExpand, browser, timeout);
			System.out.println("failed");
			sheet.getRow(i).createCell(3).setCellValue("Failed");
			sheet.getRow(i).createCell(4).setCellValue("late charge calculation method");
			Updatefile(f, wb);
		}
	
		
		
		}
		catch(Exception e){
			System.out.println("button not found");
			WebElement saveAndClose = browser.findElement(By.xpath("//button[text()='ave and Close']"));
			waitUntilElementClickable("saveAndClose", saveAndClose, browser, timeout);
			Thread.sleep(3000);
			WebElement searchExpand = browser.findElement(By.xpath("//*[contains(@id,'AP1:q1::_afrDscl')]"));
			waitUntilElementClickable("searchExpand", searchExpand, browser, timeout);
			sheet.getRow(i).createCell(3).setCellValue("Failed");
			//sheet.getRow(i).createCell(4).setCellValue("late charge calculation method");
			Updatefile(f, wb);
		}
		}
		try
		{
			wb.close();
		}
		catch(Exception e)
		{
			
		}
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
	
}
