package com.Boyd.ManageTrancations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

public class EmailUpdation extends Base_Class{
	
		public String organisationName;
		public String contactName;
		public String siteNumber;
		@Test
		public void manageCustomers() throws Exception{
			File f=new File(System.getProperty("user.dir")+"\\Excel\\EmailUpdation.xlsx");
			FileInputStream fis=new FileInputStream(f);
			XSSFWorkbook wb=new XSSFWorkbook(fis);
			XSSFSheet sheet=wb.getSheet("Sheet1");
			int totalrows = sheet.getPhysicalNumberOfRows();
			WebElement tasks = browser.findElement(By.xpath("//img[contains(@title,'Tasks')]"));
			waitUntilElementClickable("tasks", tasks, browser, timeout);
			Thread.sleep(13000);
			WebElement ManageCustomers = browser.findElement(By.xpath("//a[text()='Manage Customers']"));
			waitUntilElementClickable("ManageCustomers", ManageCustomers, browser, timeout);
			for(int i=1;i<totalrows;i++)
			{
			System.out.println("Count of i value :" +i);
			DataFormatter df = new DataFormatter();
			organisationName = df.formatCellValue(sheet.getRow(i).getCell(0));
			System.out.println(organisationName);
			DataFormatter df1 = new DataFormatter();
			siteNumber=df1.formatCellValue(sheet.getRow(i).getCell(1));
			System.out.println(siteNumber);
			contactName=sheet.getRow(i).getCell(2).getStringCellValue();
			String firstname=contactName.split(" ")[0];
			System.out.println(firstname);
			Thread.sleep(2000);
			WebElement orgName = browser.findElement(By.xpath("//input[contains(@id,'AP1:q1:value120::content')]"));
			waitUntilElementClickable("orgName", orgName, browser, timeout);
			orgName.clear();
			orgName.sendKeys(organisationName);
			WebElement search = browser.findElement(By.xpath("//*[contains(@id,':q1::search')]"));
			waitUntilElementClickable("search", search, browser, timeout);
			Thread.sleep(4000);
			WebElement siteInput = browser.findElement(By.xpath("//*[contains(@id,'ATp_afr_table1_afr_column66::content')]"));
			((JavascriptExecutor)browser).executeScript("arguments[0].scrollIntoView(true);", new Object[] { siteInput });
			waitUntilElementClickable("siteInput", siteInput, browser, timeout);
			siteInput.clear();
			siteInput.sendKeys(siteNumber);
			siteInput.sendKeys(Keys.ENTER);
			Thread.sleep(2000);
			try {
			try {	
			WebElement siteText = browser.findElement(By.xpath("//a[text()='" + siteNumber + "']"));
			waitUntilElementClickable("siteText", siteText, browser, timeout);}
			catch (Exception e){
				WebElement view = browser.findElement(By.xpath("//div[contains(@id,'1:cupt1:CManF:0:AP1:pt_r3:0:AT1:_ATp:_vw')]//a[text()='View']"));
				waitUntilElementClickable("view", view, browser, timeout);
				WebElement All = browser.findElement(By.xpath("//tr[contains(@id,'AP1:pt_r3:0:AT1:_ATp:ViewCurrentAccounts')]/..//td[text()='All']"));
				waitUntilElementClickable("All", All, browser, timeout);
				Thread.sleep(8000);
				WebElement siteInput2 = browser.findElement(By.xpath("//*[contains(@id,'ATp_afr_table1_afr_column66::content')]"));
			//	((JavascriptExecutor)browser).executeScript("arguments[0].scrollIntoView(true);", new Object[] { siteInput });
				waitUntilElementClickable("siteInput", siteInput2, browser, timeout);
				siteInput2.clear();
				siteInput2.sendKeys(siteNumber);
				siteInput2.sendKeys(Keys.ENTER);
				Thread.sleep(5000);
				WebElement siteText = browser.findElement(By.xpath("//a[text()='" + siteNumber + "']"));
				waitUntilElementClickable("siteText", siteText, browser, timeout);
				}
			WebElement communiucation = browser.findElement(By.xpath("//a[@id='pt1:_FOr1:1:_FOSritemNode_receivables_receivables_balances:0:MAnt2:1:cupt1:CManF:1:cupanel1:cushowDetailItem1::disAcr']"));
			waitUntilElementClickable("communiucation", communiucation, browser, timeout);
			//Thread.sleep(4000);
			try {
				WebElement editcontacts = browser.findElement(By.xpath("//button[text()='Edit Contacts']"));
				((JavascriptExecutor)browser).executeScript("arguments[0].scrollIntoView(true);", new Object[] { editcontacts });
				waitUntilElementClickable("editcontacts", editcontacts, browser, timeout);
				//Thread.sleep(6000);
			WebElement responsibilities = browser.findElement(By.xpath("//h2[text()='Responsibilities']"));
			responsibilities.isDisplayed();
			WebElement FirstName = browser.findElement(By.xpath("//a[contains(text(),'"+firstname+"')]"));
			waitUntilElementClickable("FirstName", FirstName, browser, timeout);
			//Thread.sleep(5000);
			try {
			WebElement emailType = browser.findElement(By.xpath("//a[translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = 'receivables@aavid.com']/../../..//span[text()='E-mail']"));
			waitUntilElementClickable("emailType", emailType, browser, timeout);
			Thread.sleep(1000);}
			catch (Exception e) {
				System.out.println("Email already updated");
				WebElement emailType1 = browser.findElement(By.xpath("//span[text()='E-mail']/../..//a[text()='thermal_receivables@boydcorp.com']"));
				System.out.println(emailType1.getText());
				WebElement emailType2 = browser.findElement(By.xpath("//a[text()='thermal_receivables@boydcorp.com']/../../..//span[text()='E-mail']"));
				waitUntilElementClickable("emailType", emailType2, browser, timeout);
				Thread.sleep(1000);
			}
			WebElement edit = browser.findElement(By.xpath("//img[@id='pt1:_FOr1:1:_FOSritemNode_receivables_receivables_balances:0:MAnt2:1:cupt1:CManF:2:AP1:region1:0:AT1:_ATp:edit::icon']"));
			waitUntilElementClickable("edit", edit, browser, timeout);
			//Thread.sleep(1000);
			WebElement inputEmail = browser.findElement(By.xpath("//input[contains(@id,'pt_r1:1:inputText2::content')]"));
			waitUntilElementClickable("inputEmail", inputEmail, browser, timeout);
			inputEmail.clear();
			Thread.sleep(1000);
			inputEmail.sendKeys("thermal_receivables@boydcorp.com");
			Thread.sleep(1000);
			WebElement okButton = browser.findElement(By.xpath("//button[text()='O']"));
			waitUntilElementClickable("okButton", okButton, browser, timeout);
			Thread.sleep(1000);
			WebElement saveAndClose = browser.findElement(By.xpath("//button[text()='ave and Close']"));
			waitUntilElementClickable("saveAndClose", saveAndClose, browser, timeout);
			//Thread.sleep(4000);
			WebElement saveAndClose1 = browser.findElement(By.xpath("//button[contains(@id,'MAnt2:1:cupt1:CManF:1:cupanel1:cucommandButton2')]"));
			waitUntilElementClickable("saveAndClose1", saveAndClose1, browser, timeout);
			Thread.sleep(3000);
			System.out.println("=====updated=====");
			WebElement searchExpand = browser.findElement(By.xpath("//*[contains(@id,'AP1:q1::_afrDscl')]"));
			waitUntilElementClickable("searchExpand", searchExpand, browser, timeout);
			sheet.getRow(i).createCell(3).setCellValue("updated");
			Updatefile(f, wb);
			//Thread.sleep(3000);
			}
			catch (Exception e){
				System.out.println("=====Failed=====");
				System.out.println("=====contact=====");
				sheet.getRow(i).createCell(3).setCellValue("Failed");
				sheet.getRow(i).createCell(4).setCellValue("Contact or Email mismatch");
				Updatefile(f, wb);
				WebElement saveAndClose3 = browser.findElement(By.xpath("//button[text()='ave and Close']"));
				waitUntilElementClickable("saveAndClose", saveAndClose3, browser, timeout);
				Thread.sleep(4000);
				WebElement saveAndClose2 = browser.findElement(By.xpath("//button[contains(@id,'MAnt2:1:cupt1:CManF:1:cupanel1:cucommandButton2')]"));
				waitUntilElementClickable("saveAndClose1", saveAndClose2, browser, timeout);
				//Thread.sleep(5000);
				WebElement searchExpand = browser.findElement(By.xpath("//*[contains(@id,'AP1:q1::_afrDscl')]"));
				waitUntilElementClickable("searchExpand", searchExpand, browser, timeout);
				
			}
			}	
			catch(Exception e)
			{
				e.printStackTrace();
				System.out.println("=====Failed=====");
				sheet.getRow(i).createCell(3).setCellValue("Failed");
				sheet.getRow(i).createCell(4).setCellValue("site not found");
				Updatefile(f, wb);
				
				WebElement searchExpand = browser.findElement(By.xpath("//*[contains(@id,'AP1:q1::_afrDscl')]"));
				waitUntilElementClickable("searchExpand", searchExpand, browser, timeout);
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