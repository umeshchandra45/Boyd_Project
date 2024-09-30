package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BOYD_P2P_Payments_to_Supplier {
	public WebDriver browser;
	public String Name;
	public String Template;
	public String Pay_From_Date;
	public String Invoice_group;
	public String Priority_Override;
	public String Comments;
	public String Bank_Account;
	public String Business_Unit;
	public String Legal_Entity;
	public String Payment_Process_request;
	public String Payment_Document;
	public String Transmit_Now;
	
	@BeforeTest()
	public void Login_Page() throws Exception
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
		browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
	//	browser.get("https://elme-dev2.fa.us8.oraclecloud.com");
	//	browser.get("https://elme-test.login.us8.oraclecloud.com/");
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("Laura.kelly");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("Welcome1");
		browser.findElement(By.id("btnActive")).click();
		Thread.sleep(25000);
		browser.findElement(By.xpath("//a[text()='You have a new home page!']")).click();
		Thread.sleep(12000);
		browser.findElement(By.linkText("Payables")).click();
		WebElement payments = browser.findElement(By.linkText("Payments"));
		WebDriverwaitelement(payments);
		payments.click();
	}
	
	@SuppressWarnings("unused")
	@Test()
	public void Home_Page() throws Exception
	{
		
		File f = new File(System.getProperty("user.dir")+"\\Excel\\BOYD_P2P_Payments_to_Supplier.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Payments_to_Supplier");
		sheet.getRow(0).createCell(12).setCellValue("Result");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		
	//	if(sheet.getRow(1).getCell(12) == null)
	//	{
		
		for(int i=1;i<=totalrows;i++)
		{
			if(sheet.getRow(i) == null)
			{
				return;
			}
			
			Name = sheet.getRow(i).getCell(0).getStringCellValue();
			Template = sheet.getRow(i).getCell(1).getStringCellValue();
			Pay_From_Date = sheet.getRow(i).getCell(2).getStringCellValue();
			Invoice_group = sheet.getRow(i).getCell(3).getStringCellValue();
			Priority_Override = sheet.getRow(i).getCell(4).getStringCellValue();
			Comments = sheet.getRow(i).getCell(5).getStringCellValue();
			Bank_Account = sheet.getRow(i).getCell(6).getStringCellValue();
			Business_Unit = sheet.getRow(i).getCell(7).getStringCellValue();
			Legal_Entity = sheet.getRow(i).getCell(8).getStringCellValue();
			Payment_Process_request = sheet.getRow(i).getCell(9).getStringCellValue();
			Payment_Document = sheet.getRow(i).getCell(10).getStringCellValue();
			Transmit_Now = sheet.getRow(i).getCell(11).getStringCellValue();
			
			
			Thread.sleep(6000);
			WebElement task = browser.findElement(By.xpath("//*[contains(@id,'_FOTsdi__PaymentLanding_itemNode__FndTasksList::icon')]"));
			WebDriverwaitelement(task);
			task.click();
			WebElement paymentprocess = browser.findElement(By.linkText("Submit Payment Process Request"));
			WebDriverwaitelement(paymentprocess);
			paymentprocess.click();
			WebElement name = browser.findElement(By.xpath("//*[contains(@id,'inputText1::content')]"));
			WebDriverwaitelement(name);
			name.click();
			name.sendKeys(Name);
			Thread.sleep(4000);
			browser.findElement(By.xpath("//*[contains(@id,'templateNameId::lovIconId')]")).click();
			WebElement searchicon = browser.findElement(By.linkText("Search..."));
			WebDriverwaitelement(searchicon);
			searchicon.click();
			WebElement namefield = browser.findElement(By.xpath("//*[contains(@id,'templateNameId::_afrLovInternalQueryId:value00::content')]"));
			WebDriverwaitelement(namefield);
			namefield.click();
			namefield.clear();
			namefield.sendKeys(Template);
			Thread.sleep(2000);
			browser.findElement(By.xpath("//button[text()='Search']")).click();
			WebElement tablerow = browser.findElement(By.xpath("//*[contains(@id,'templateNameId_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
			WebDriverwaitelement(tablerow);
			tablerow.click();
			Thread.sleep(4000);
			browser.findElement(By.xpath("//*[contains(@id,'templateNameId::lovDialogId::ok')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("//*[contains(@id,'inputDate4::content')]")).click();
			browser.findElement(By.xpath("//*[contains(@id,'inputDate4::content')]")).sendKeys(Pay_From_Date);
			Thread.sleep(4000);
//			browser.findElement(By.xpath("//*[contains(@id,'batchNameId::lovIconId')]")).click();
//			WebElement invoicesearch = browser.findElement(By.linkText("Search..."));
//			WebDriverwaitelement(invoicesearch);
//			invoicesearch.click();
//			WebElement invoicegroup = browser.findElement(By.xpath("//*[contains(@id,'batchNameId::_afrLovInternalQueryId:value00::content')]"));
//			WebDriverwaitelement(invoicegroup);
//			invoicegroup.click();
//			invoicegroup.sendKeys(Invoice_group);
//			Thread.sleep(3000);
//			browser.findElement(By.xpath("//*[contains(@id,'batchNameId::_afrLovInternalQueryId::search')]")).click();
//			WebElement invoicetable = browser.findElement(By.xpath("//*[contains(@id,'batchNameId_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
//			WebDriverwaitelement(invoicetable);
//			invoicetable.click();
//			Thread.sleep(3000);
//			browser.findElement(By.xpath("//*[contains(@id,'batchNameId::lovDialogId::ok')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.linkText("Payment and Processing Options")).click();
			Thread.sleep(6000);
			WebElement el =  browser.findElement(By.xpath("//*[contains(@id,'selectBooleanCheckbox3::content')]"));
			JavascriptExecutor js = (JavascriptExecutor)browser;
			js.executeScript("arguments[0].click()", el);
			Thread.sleep(4000);
			Select po = new Select(browser.findElement(By.xpath("//*[contains(@id,'selectOneChoice7::content')]")));
			po.selectByVisibleText(Priority_Override);
			Thread.sleep(4000);
			WebElement el1 = browser.findElement(By.xpath("//*[contains(@id,'ap1:sbc1::content')]"));
			JavascriptExecutor js1 = (JavascriptExecutor)browser;
			js1.executeScript("arguments[0].click()", el1);
			Thread.sleep(4000);
			browser.findElement(By.xpath("//*[contains(@id,'SPsb2')]")).click();
			WebElement task1 = browser.findElement(By.xpath("//*[contains(@id,'_FOTsdi__PaymentLanding_itemNode__FndTasksList::icon')]"));
			WebDriverwaitelement(task1);
			task1.click();
			WebElement requesttemplate = browser.findElement(By.linkText("Manage Payment Process Requests"));
			WebDriverwaitelement(requesttemplate);
			requesttemplate.click();
			WebElement requestname = browser.findElement(By.xpath("//*[contains(@id,'q1:value00::content')]"));
			WebDriverwaitelement(requestname);
			requestname.click();
			requestname.sendKeys(Name);
			Thread.sleep(4000);
			browser.findElement(By.xpath("//*[contains(@id,'q1::search')]")).click();
			Thread.sleep(6000);
			for(int k=1;k<=5;k++)
			{
				browser.findElement(By.xpath("//*[contains(@id,'ATT2:_ATTp:tti2::icon')]")).click();
				Thread.sleep(6000);
			}
			//Thread.sleep(6000);
			browser.findElement(By.xpath("//*[contains(@id,'ATT2:_ATTp:tt1::db')]/table/tbody/tr/td[1]")).click();
			Thread.sleep(7000);
			browser.findElement(By.xpath("//*[contains(@id,'commandImageLink2_1::icon')]")).click();
			Thread.sleep(4000);
			WebElement submit = browser.findElement(By.xpath("//*[contains(@id,'AP1:SPsb2')]"));
			WebDriverwaitelement(submit);
			submit.click();
			Thread.sleep(6000);
			for(int k=1;k<=7;k++)
			{
				browser.findElement(By.xpath("//*[contains(@id,'_ATTp:ATTtb2::oc')]")).click();
				Thread.sleep(8000);
			}
			Thread.sleep(6000);
			WebElement actionbutton = browser.findElement(By.xpath("//*[contains(@id,'commandImageLink2_1::icon')]"));
			WebDriverwaitelement(actionbutton);
			actionbutton.click();
			WebElement resumepaymentprocess = browser.findElement(By.xpath("//button[text()='Resume Payment Process']"));
			WebDriverwaitelement(resumepaymentprocess);
			resumepaymentprocess.click();
			Thread.sleep(8000);
			browser.findElement(By.id("pt1:_UISatr:0:cil1::icon")).click();
			WebElement moredetails = browser.findElement(By.linkText("More Details"));
			WebDriverwaitelement(moredetails);
			moredetails.click();
			Thread.sleep(10000);
			String mainwindow = browser.getWindowHandle();
			Set<String> windows = browser.getWindowHandles();
			System.out.println("number of windows :" +windows.size());
			String [] array = windows.toArray(new String[windows.size()]);
			String window1 = array[0];
			String window2 = array[1];
			browser.switchTo().window(window2);
			Thread.sleep(15000);
			browser.findElement(By.xpath("//a[contains(@id,'vinp1:j_id__ctru5pc14j_id_1')]")).click();
//			Thread.sleep(3000);
//			browser.findElement(By.xpath("//*[contains(@id,'dc_np1:wvemij_id_1')]")).click();
			Thread.sleep(6000);
			WebElement approval = browser.findElement(By.xpath("//a[text()='Approval of Payment Process Request "+Name+"']"));
			System.out.println("Approval value is :" +approval);
			approval.click();
			Thread.sleep(8000);
			Set<String> subwindow = browser.getWindowHandles();
			System.out.println("Number of windows :" +subwindow.size());
			String[] subarray = subwindow.toArray(new String[subwindow.size()]);
			String subwindow1 = subarray[0];
			String subwindow2 = subarray[1];
			String subwindow3 = subarray[2];
			browser.switchTo().window(subwindow3);
			Thread.sleep(8000);
			WebElement subactionwindow = browser.findElement(By.linkText("Actions"));
			WebDriverwaitelement(subactionwindow);
			subactionwindow.click();
			Thread.sleep(4000);
			browser.findElement(By.xpath("//td[text()='View Approvals']")).click();
			WebElement okbutton = browser.findElement(By.id("r1:0:bip_up:UPsp1:bip_rpp:0:ctb_aph_Ok"));
			WebDriverwaitelement(okbutton);
			okbutton.click();
			Thread.sleep(4000);
			browser.close();
			browser.switchTo().window(subwindow2);
			Thread.sleep(4000);
			browser.close();
			Thread.sleep(4000);
			browser.switchTo().window(window1);
			Thread.sleep(6000);
			browser.findElement(By.id("pt1:_UIScmil1u::icon")).click();
			WebElement signout = browser.findElement(By.linkText("Sign Out"));
			WebDriverwaitelement(signout);
			signout.click();
			WebElement confirmbutton = browser.findElement(By.id("Confirm"));
			WebDriverwaitelement(confirmbutton);
			confirmbutton.click();
			WebElement username = browser.findElement(By.id("userid"));
			WebDriverwaitelement(username);
			username.click();
			username.sendKeys("jerry.bellerose");
			WebElement password = browser.findElement(By.id("password"));
			WebDriverwaitelement(password);
			password.click();
			password.sendKeys("Welcome1");
			Thread.sleep(3000);
			browser.findElement(By.id("btnActive")).click();
			Thread.sleep(10000);
			browser.findElement(By.id("pt1:_UISatr:0:cil1::icon")).click();
			WebElement moredetails1 = browser.findElement(By.linkText("More Details"));
			WebDriverwaitelement(moredetails1);
			moredetails1.click();
			Thread.sleep(6000);
			Set<String> windows1 = browser.getWindowHandles();
			System.out.println("Number of windows :" +windows1.size());
			String[] arraywindow = windows1.toArray(new String[windows1.size()]);
			String subwind = arraywindow[0];
			String subwind1 = arraywindow[1];
			browser.switchTo().window(subwind1);
			Thread.sleep(10000);
			browser.findElement(By.xpath("//a[text()='Approval of Payment Process Request "+Name+"']")).click();
			Thread.sleep(4000);
			Set<String> windows2 = browser.getWindowHandles();
			System.out.println("Number of windows :" +windows2.size());
			String[] arraywindow1 = windows2.toArray(new String[windows2.size()]);
			String subwind2 = arraywindow1[0];
			String subwind3 = arraywindow1[1];
			String subwind4 = arraywindow1[2];
			browser.switchTo().window(subwind4);
			Thread.sleep(6000);
			browser.findElement(By.xpath("//button[text()='Approve']")).click();
			Thread.sleep(8000);
			browser.findElement(By.xpath("//*[contains(@id,'it_apprej::content')]")).click();
			browser.findElement(By.xpath("//*[contains(@id,'it_apprej::content')]")).sendKeys(Comments);
			Thread.sleep(4000);
			browser.findElement(By.xpath("//span[text()='Submit']")).click();
			Thread.sleep(6000);
			browser.switchTo().window(subwind1);
			Thread.sleep(3000);
			browser.close();
			Thread.sleep(6000);
			browser.switchTo().window(subwind);
			Thread.sleep(6000);
			browser.findElement(By.id("pt1:_UIScmil1u::icon")).click();
			WebElement sign = browser.findElement(By.linkText("Sign Out"));
			WebDriverwaitelement(sign);
			sign.click();
			WebElement con = browser.findElement(By.id("Confirm"));
			WebDriverwaitelement(con);
			con.click();
			WebElement user = browser.findElement(By.id("userid"));
			WebDriverwaitelement(user);
			user.click();
			user.sendKeys("Laura.kelly");
			Thread.sleep(3000);
			browser.findElement(By.id("password")).click();
			browser.findElement(By.id("password")).sendKeys("Welcome1");
			Thread.sleep(3000);
			browser.findElement(By.id("btnActive")).click();
			Thread.sleep(12000);
			browser.findElement(By.linkText("Payables")).click();
			Thread.sleep(2000);
			browser.findElement(By.linkText("Payments")).click();
			Thread.sleep(6000);
//			browser.findElement(By.xpath("//*[text()='"+Name+"']/../../../../..//*[contains(@id,'commandImageLink1_1::icon')]")).click();
			WebElement actionbutton1 = browser.findElement(By.xpath("//*[text()='"+Name+"']/../../../../..//*[contains(@id,'commandImageLink1_1::icon')]"));
			JavascriptExecutor js5 = (JavascriptExecutor)browser;
			js5.executeScript("arguments[0].scrollIntoView();", actionbutton1);
			actionbutton1.click();
			WebElement resumepayment = browser.findElement(By.xpath("//button[text()='Resume Payment Process']"));
			WebDriverwaitelement(resumepayment);
			resumepayment.click();
			WebElement taskicon = browser.findElement(By.xpath("//*[contains(@id,'_FOTsdi__PaymentLanding_itemNode__FndTasksList::icon')]"));
			WebDriverwaitelement(taskicon);
			taskicon.click();
			WebElement electronicpayment = browser.findElement(By.linkText("Create Electronic Payment Files"));
			WebDriverwaitelement(electronicpayment);
			electronicpayment.click();
			WebElement bankaccount = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BankAccountName::lovIconId')]"));
			WebDriverwaitelement(bankaccount);
			bankaccount.click();
			WebElement search = browser.findElement(By.linkText("Search..."));
			WebDriverwaitelement(search);
			search.click();
			WebElement account = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BankAccountName::_afrLovInternalQueryId:value00::content')]"));
			WebDriverwaitelement(account);
			account.click();
			account.sendKeys(Bank_Account);
			Thread.sleep(2000);
			browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BankAccountName::_afrLovInternalQueryId::search')]")).click();
			WebElement tablerow1 = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BankAccountName_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
			WebDriverwaitelement(tablerow1);
			tablerow1.click();
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BankAccountName::lovDialogId::ok')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BusinessUnit::lovIconId')]")).click();
			WebElement BU = browser.findElement(By.linkText("Search..."));
		    WebDriverwaitelement(BU);
		    BU.click();
		    WebElement Bussinessunit = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BusinessUnit::_afrLovInternalQueryId:value00::content')]"));
		    WebDriverwaitelement(Bussinessunit);
		    Bussinessunit.click();
		    Bussinessunit.sendKeys(Business_Unit);
		    Thread.sleep(2000);
		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BusinessUnit::_afrLovInternalQueryId::search')]")).click();
		    WebElement BUtable = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BusinessUnit_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
		    WebDriverwaitelement(BUtable);
		    BUtable.click();
		    Thread.sleep(3000);
		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_BusinessUnit::lovDialogId::ok')]")).click();
		    Thread.sleep(6000);
		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_LegalEntity::lovIconId')]")).click();
		    WebElement LE = browser.findElement(By.linkText("Search..."));
		    WebDriverwaitelement(LE);
		    LE.click();
		    WebElement Legalentityname = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_LegalEntity::_afrLovInternalQueryId:value00::content')]"));
		    WebDriverwaitelement(Legalentityname);
		    Legalentityname.click();
		    Legalentityname.sendKeys(Legal_Entity);
		    Thread.sleep(3000);
		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_LegalEntity::_afrLovInternalQueryId::search')]")).click();
		    WebElement letable = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_LegalEntity_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
		    WebDriverwaitelement(letable);
		    letable.click();
		    Thread.sleep(3000);
		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_LegalEntity::lovDialogId::ok')]")).click();
		    Thread.sleep(6000);
//		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute12_ATTRIBUTE12::lovIconId')]")).click();
//		    WebElement paymenttable = browser.findElement(By.linkText("Search..."));
//		    WebDriverwaitelement(paymenttable);
//		    paymenttable.click();
//		    WebElement paymentrequest = browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute12_ATTRIBUTE12::_afrLovInternalQueryId:value00::content')]"));
//		    WebDriverwaitelement(paymentrequest);
//		    paymentrequest.click();
//		    paymentrequest.sendKeys(Payment_Process_request);
//		    Thread.sleep(3000);
//		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute12_ATTRIBUTE12::_afrLovInternalQueryId::search')]")).click();
//		    Thread.sleep(6000);
//		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute12_ATTRIBUTE12_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
//		    Thread.sleep(3000);
//		    browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute12_ATTRIBUTE12::lovDialogId::ok')]")).click();
//		    Thread.sleep(6000);
		    Select paymentdocument = new Select(browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute3_ATTRIBUTE3::content')]")));
		    paymentdocument.selectByVisibleText(Payment_Document);
		    Thread.sleep(4000);
		    Select transmit = new Select(browser.findElement(By.xpath("//*[contains(@id,'basicReqBody:paramDynForm_Attribute8_ATTRIBUTE8::content')]")));
		    transmit.selectByVisibleText(Transmit_Now);
		    Thread.sleep(6000);
		    browser.findElement(By.xpath("//*[contains(@id,'requestBtns:submitButton')]")).click();
		    Thread.sleep(8000);
		    browser.findElement(By.xpath("//*[contains(@id,'requestBtns:confirmationPopup:confirmSubmitDialog::ok')]")).click();
		    Thread.sleep(6000);
		    browser.findElement(By.id("pt1:_UIShome")).click();
		    Thread.sleep(8000);
		    browser.findElement(By.linkText("Payables")).click();
			Thread.sleep(2000);
			browser.findElement(By.linkText("Payments")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("(//*[contains(@id,'sdi2::disAcr')])[1]")).click();
			Thread.sleep(4000);
			browser.findElement(By.xpath("//*[contains(@id,'RecentlyCompletedPpr:_ATTp:tt2::db')]/table/tbody/tr[1]/td[1]")).click();
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'RecentlyCompletedPpr:_ATTp:tt2:0::di')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("(//*[contains(@id,'commandLink3')])[2]")).click();
			Thread.sleep(10000);
			JavascriptExecutor js3 = (JavascriptExecutor)browser;
			js3.executeScript("window.scrollBy(0,450)");
			Thread.sleep(3000);
			browser.findElement(By.xpath("//*[contains(@id,'commandImageLink1::icon')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("//*[contains(@id,'commandButton1')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("//*[contains(@id,'ap1:commandButton1')]")).click();
			Thread.sleep(6000);
			browser.findElement(By.xpath("(//*[contains(@id,'sdi1::disAcr')])[1]")).click();
		    sheet.getRow(i).createCell(12).setCellValue("Pass");
			Updatefile(f, wb);
		}
//	}
		/*else
		{
			System.out.println("File is already Processed");
		}*/
		try
		{
			wb.close();
		}
		catch(Exception e)
		{
			
		}
		
		
	}
	
	public void WebDriverwaitelement(WebElement element)
	{
		WebDriverWait wait = new WebDriverWait(browser,350);
		wait.until(ExpectedConditions.visibilityOf(element));
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
	
	@AfterTest()
	public void Close_Browser()
	{
//		browser.quit();
	}
	
	

}
