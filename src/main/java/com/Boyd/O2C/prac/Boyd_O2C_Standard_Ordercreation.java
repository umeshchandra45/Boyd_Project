package com.Boyd.O2C.prac;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
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

		@SuppressWarnings("unused")
		public class Boyd_O2C_Standard_Ordercreation {
		public WebDriver browser;
		public String Business_Unit;
		public String Customer;
		public String PurchaseOrder;
		public String OrderType;
		public String Contact;
		public String Contact_Method;
		public String Ship_to_Address;
		public String Bill_to_Customer;
		public String Bill_to_Address;
		public String Manage_Attachments;
		public String EndCustomer_Name;
		public String Sales_Order_Acknowledgement_required;
		public String Header_Notes;
		public String Quality_Rating;
		public String BDE;
		public String Sub_End_Customer;
		public String DPAS_Agency;
		public String DPAS_Program_ID;
		public String DPAS_Rating;
		public String Government_Contract;
		public String ITAR_Restricted;
		public String Export_License_Reqd;
		public String FARS;
		public String DFARS;
		public String Group;
		public String Region;
		public String Region2;
		public String Ship_To_Contact;
		public String Ship_to_Contact_Method;
		public String Bill_To_Contact;
		public String Customer_Email_Address;
		public String PhoneNumber;
		public String Payment_Terms;
		public String Shipment_Priority_Header;
		public String ShippingMethod;
		public String Requested_Date_Header;
		public String Request_Type_Header;
		public String FOB_Header;
		public String FrieghtTerms_Header;
		public String Shipping_Instructions_Header;
		public String Packing_Instructions_header;
		public String Allow_partials;
		public String Warehouse_Header;
		public String Demand_Class_Header;
		public String Supplier_Header;
		public String PricingSubEnd_Customer;
		public String Program;
		public String Platform;
		public String TradingPartnerItem;
		public String PrimarySalesperson;
		
		public String SubIndustry_Segment;
		public String Item;
		public String SubEnd_Customer;
		public String Quantity;
		public String Adjustment_Type;
		public String Unit_Selling_Price;
		public String Reason;
		public String Manage_Attachements2;
		public String MTN_field;
		public String Repromise_Date;
		public String Original_Schedule_Ship_Date;
		public String Customer_Catalouge_Cross_Refrerence;
		public String Additional_Notes;
		public String MPN;
		public String Internal_Item;
		public String Cust_Src_Inspec;
		public String Govt_Src_Inspec;
		public String FAI;
		public String Material_Certs;
		public String Test_Reports;
		public String Dimensional_Inspection;
		public String FAA_form;
		public String FOB;
		public String FrieghtTerms;
		public String Shipping_Instructions;
		public String Packing_Instructions;
		public String ShippingMethod_Header;
		public String Requested_Date;
		public String Request_Type;
		public String Warehouse;
		public String Demand_Class;
		public String Purchase_Order_Number;
		public String Purchase_Order_Line;
		public String Manage_Attachments1;
		public static WebDriverWait wait;
		public static int timeout = 40;
		
		
		
		
		
		
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

		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("forsys2023");
		browser.findElement(By.id("btnActive")).click();
		WebElement homebutton = browser.findElement(By.id("pt1:_UIShome"));
		waitUntilElementClickable("homebutton", homebutton, browser, timeout);
		WebElement ordermanagement = browser.findElement(By.linkText("Order Management"));
		waitUntilElementClickable("ordermanagement", ordermanagement, browser, timeout);
		WebElement order1 = browser.findElement(By.id("itemNode_order_management_order_management_1"));
		waitUntilElementClickable("order1", order1, browser, timeout);
		WebElement create = browser.findElement(By.xpath("//span[text()='Create Order']"));
		waitUntilElementClickable("create", create, browser, timeout);
		}
		
		@Test() 
		public void Home_Page() throws Exception
		{
		
		File f = new File(System.getProperty("user.dir")+"\\Excel\\BOYD_O2C_Sales_Order_Creation1.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Salesordercreation");
		sheet.getRow(0).createCell(84).setCellValue("Order");
		sheet.getRow(0).createCell(85).setCellValue("Result");
		sheet.getRow(0).createCell(86).setCellValue("Comments");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		
		if(sheet.getRow(1).getCell(85) == null || sheet.getRow(1).getCell(80).getStringCellValue().contentEquals(""))
		{
		
		for(int i=1;i<=totalrows;i++)
		{
		if(sheet.getRow(i) == null)
		{
		try
		{
		WebElement save = browser.findElement(By.xpath("//*[text()='Save']"));
		waitUntilElementClickable("save", save, browser, timeout);
		Thread.sleep(14000);
		WebElement status = browser.findElement(By.xpath("//*[contains(@id,'SPph::_afrTtxt')]"));
		Thread.sleep(3000);
		String orderstatus = status.getText();
		Thread.sleep(3000);
		int ordernumber = getNumericValue(orderstatus);
		System.out.println("Order number is :" +ordernumber);
		Thread.sleep(10000);
		WebElement submit = browser.findElement(By.xpath("//span[text()='Submit']"));
		waitUntilElementClickable("submit", submit, browser, timeout);
		WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'cb14')]"));
		waitUntilElementClickable("okbutton", okbutton, browser, timeout);
		WebElement done = browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]"));
		waitUntilElementClickable("done", done, browser, timeout);
		sheet.getRow(i-1).createCell(84).setCellValue(ordernumber);
		Updatefile(f, wb);
		break;
		}
		catch(Exception e)
		{
		WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'d4::ok')]"));
		waitUntilElementClickable("okbutton", okbutton, browser, timeout);
		Thread.sleep(6000);
		WebElement cancelbutton = browser.findElement(By.xpath("(//*[text()='ancel'])[1]"));
		waitUntilElementClickable("cancelbutton", cancelbutton, browser, timeout);
		WebElement waring = browser.findElement(By.xpath("(//*[contains(@id,'cb4')])[2]"));
		waitUntilElementClickable("waring", waring, browser, timeout);
		WebElement done = browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]"));
		waitUntilElementClickable("done", done, browser, timeout);
		sheet.getRow(i-1).createCell(85).setCellValue("Fail");
		sheet.getRow(i-1).createCell(86).setCellValue("The value provided for the attribute Salesperson is invalid");
		Updatefile(f, wb);
		break;
		}
		}
		
		if(sheet.getRow(i) != null && isRowEmpty(sheet.getRow(i)))
		{
		try {
		WebElement savebutton = browser.findElement(By.xpath("//*[text()='Save']"));
		waitUntilElementClickable("savebutton", savebutton, browser, timeout);
		Thread.sleep(14000);
		WebElement status = browser.findElement(By.xpath("//*[contains(@id,'SPph::_afrTtxt')]"));
		String orderstatus = status.getText();
		int ordernumber = getNumericValue(orderstatus);
		System.out.println("Order number is :" +ordernumber);
		Thread.sleep(10000);
		WebElement submitbutton = browser.findElement(By.xpath("//span[text()='Submit']"));
		waitUntilElementClickable("submitbutton", submitbutton, browser, timeout);
		WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'cb14')]"));
		waitUntilElementClickable("okbutton", okbutton, browser, timeout);
		WebElement done = browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]"));
		waitUntilElementClickable("done", done, browser, timeout);
		WebElement createorder = browser.findElement(By.xpath("//span[text()='Create Order']"));
		waitUntilElementClickable("createorder", createorder, browser, timeout);
		sheet.getRow(i-1).createCell(77).setCellValue(ordernumber);
		Updatefile(f, wb);
		continue;
		}
		catch(Exception e)
		{
		WebElement okbutton1 = browser.findElement(By.xpath("//*[contains(@id,'d4::ok')]"));
		waitUntilElementClickable("okbutton1", okbutton1, browser, timeout);
		browser.findElement(By.xpath("(//*[text()='ancel'])[1]")).click();
		WebElement waring = browser.findElement(By.xpath("(//*[contains(@id,'cb4')])[2]"));
		waitUntilElementClickable("waring", waring, browser, timeout);
		WebElement done = browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]"));
		waitUntilElementClickable("done", done, browser, timeout);
		WebElement createorder = browser.findElement(By.xpath("//span[text()='Create Order']"));
		waitUntilElementClickable("createorder", createorder, browser, timeout);
		sheet.getRow(i-1).createCell(78).setCellValue("Fail");
		sheet.getRow(i-1).createCell(79).setCellValue("The value provided for the attribute Salesperson is invalid");
		Updatefile(f, wb);
		continue;
		}
		
		}
		
		
		if(sheet.getRow(i) == null)
		 {
		 return;
		 }
		
		Business_Unit = sheet.getRow(i).getCell(0).getStringCellValue();
		Customer = sheet.getRow(i).getCell(1).getStringCellValue();
		PurchaseOrder = sheet.getRow(i).getCell(2).getStringCellValue();
		OrderType = sheet.getRow(i).getCell(3).getStringCellValue();
		Contact = sheet.getRow(i).getCell(4).getStringCellValue();
		Contact_Method = sheet.getRow(i).getCell(5).getStringCellValue();
		Ship_to_Address = sheet.getRow(i).getCell(6).getStringCellValue();
		Bill_to_Customer = sheet.getRow(i).getCell(7).getStringCellValue();
		Bill_to_Address = sheet.getRow(i).getCell(8).getStringCellValue();
		Manage_Attachments = sheet.getRow(i).getCell(9).getStringCellValue();
		EndCustomer_Name = sheet.getRow(i).getCell(10).getStringCellValue();
		Sales_Order_Acknowledgement_required = sheet.getRow(i).getCell(11).getStringCellValue();
		Header_Notes = sheet.getRow(i).getCell(12).getStringCellValue();
		Quality_Rating = sheet.getRow(i).getCell(13).getStringCellValue();
		BDE = sheet.getRow(i).getCell(14).getStringCellValue();
		Sub_End_Customer = sheet.getRow(i).getCell(15).getStringCellValue();
		DPAS_Agency = sheet.getRow(i).getCell(16).getStringCellValue();
		DPAS_Program_ID = sheet.getRow(i).getCell(17).getStringCellValue();
		DPAS_Rating = sheet.getRow(i).getCell(18).getStringCellValue();
		Government_Contract = sheet.getRow(i).getCell(19).getStringCellValue();
		ITAR_Restricted = sheet.getRow(i).getCell(20).getStringCellValue();
		Export_License_Reqd = sheet.getRow(i).getCell(21).getStringCellValue();
		FARS = sheet.getRow(i).getCell(22).getStringCellValue();
		DFARS = sheet.getRow(i).getCell(23).getStringCellValue();
		Group = sheet.getRow(i).getCell(24).getStringCellValue();
		Region = sheet.getRow(i).getCell(25).getStringCellValue();
		Region2 = sheet.getRow(i).getCell(26).getStringCellValue();
		Ship_To_Contact = sheet.getRow(i).getCell(27).getStringCellValue();
		Ship_to_Contact_Method = sheet.getRow(i).getCell(28).getStringCellValue();
		Bill_To_Contact = sheet.getRow(i).getCell(29).getStringCellValue();
		Customer_Email_Address = sheet.getRow(i).getCell(30).getStringCellValue();
		PhoneNumber = sheet.getRow(i).getCell(31).getStringCellValue();
		Payment_Terms = sheet.getRow(i).getCell(32).getStringCellValue();
		Shipment_Priority_Header = sheet.getRow(i).getCell(33).getStringCellValue();
		ShippingMethod = sheet.getRow(i).getCell(34).getStringCellValue();
		Requested_Date_Header = sheet.getRow(i).getCell(35).getStringCellValue();
		Request_Type_Header = sheet.getRow(i).getCell(36).getStringCellValue();
		FOB_Header = sheet.getRow(i).getCell(37).getStringCellValue();
		FrieghtTerms_Header = sheet.getRow(i).getCell(38).getStringCellValue();
		Shipping_Instructions_Header = sheet.getRow(i).getCell(39).getStringCellValue();
		Packing_Instructions_header = sheet.getRow(i).getCell(40).getStringCellValue();
		Allow_partials = sheet.getRow(i).getCell(41).getStringCellValue();
		Warehouse_Header = sheet.getRow(i).getCell(42).getStringCellValue();
		Demand_Class_Header = sheet.getRow(i).getCell(43).getStringCellValue();
		Supplier_Header = sheet.getRow(i).getCell(44).getStringCellValue();
		
		
		Item = sheet.getRow(i).getCell(45).getStringCellValue();
		Quantity = sheet.getRow(i).getCell(46).getStringCellValue();
		Adjustment_Type = sheet.getRow(i).getCell(47).getStringCellValue();
		Unit_Selling_Price = sheet.getRow(i).getCell(48).getStringCellValue();
		Reason = sheet.getRow(i).getCell(49).getStringCellValue();
		Manage_Attachements2 = sheet.getRow(i).getCell(50).getStringCellValue();
		MTN_field = sheet.getRow(i).getCell(51).getStringCellValue();
		Repromise_Date = sheet.getRow(i).getCell(52).getStringCellValue();
		Original_Schedule_Ship_Date = sheet.getRow(i).getCell(53).getStringCellValue();
		Customer_Catalouge_Cross_Refrerence = sheet.getRow(i).getCell(54).getStringCellValue();
		Additional_Notes = sheet.getRow(i).getCell(55).getStringCellValue();
		MPN = sheet.getRow(i).getCell(56).getStringCellValue();
		Internal_Item = sheet.getRow(i).getCell(57).getStringCellValue();
		Cust_Src_Inspec = sheet.getRow(i).getCell(58).getStringCellValue();
		Govt_Src_Inspec = sheet.getRow(i).getCell(59).getStringCellValue();
		FAI = sheet.getRow(i).getCell(60).getStringCellValue();
		Material_Certs = sheet.getRow(i).getCell(61).getStringCellValue();
		Test_Reports = sheet.getRow(i).getCell(62).getStringCellValue();
		Dimensional_Inspection = sheet.getRow(i).getCell(63).getStringCellValue();
		FAA_form = sheet.getRow(i).getCell(64).getStringCellValue();
		FOB = sheet.getRow(i).getCell(65).getStringCellValue();
		FrieghtTerms = sheet.getRow(i).getCell(66).getStringCellValue();
		Shipping_Instructions = sheet.getRow(i).getCell(67).getStringCellValue();
		Packing_Instructions = sheet.getRow(i).getCell(68).getStringCellValue();
		ShippingMethod_Header = sheet.getRow(i).getCell(69).getStringCellValue();
		Requested_Date = sheet.getRow(i).getCell(70).getStringCellValue();
		Request_Type = sheet.getRow(i).getCell(71).getStringCellValue();
		Warehouse = sheet.getRow(i).getCell(72).getStringCellValue();
		Demand_Class = sheet.getRow(i).getCell(73).getStringCellValue();
		Purchase_Order_Number = sheet.getRow(i).getCell(74).getStringCellValue();
		Purchase_Order_Line = sheet.getRow(i).getCell(75).getStringCellValue();
		Manage_Attachments1 = sheet.getRow(i).getCell(76).getStringCellValue();
		SubEnd_Customer = sheet.getRow(i).getCell(77).getStringCellValue();
		PricingSubEnd_Customer=sheet.getRow(i).getCell(78).getStringCellValue();
		SubIndustry_Segment=sheet.getRow(i).getCell(79).getStringCellValue();
		Platform=sheet.getRow(i).getCell(80).getStringCellValue();
		Program=sheet.getRow(i).getCell(81).getStringCellValue();
		TradingPartnerItem=sheet.getRow(i).getCell(82).getStringCellValue();
		PrimarySalesperson=sheet.getRow(i).getCell(83).getStringCellValue();
		
		
		if(!Customer.equals("NA") && !PurchaseOrder.equals("NA"))
		{
		Headerpartupdate();
		}
		
		try
		{
			Thread.sleep(9000);
		WebElement searchicon = browser.findElement(By.xpath("//img[contains(@id, 'searchIcoId::icon')]"));
		waitUntilElementClickable("searchicon", searchicon, browser, timeout);
		WebElement advance = browser.findElement(By.xpath("//*[contains(@id,'Advan1:0:efqrp::mode')]"));
		waitUntilElementClickable("advance", advance, browser, timeout);
		Thread.sleep(6000);
		Select itemvalue = new Select(browser.findElement(By.xpath("//*[contains(@id,'efqrp:operator0::content')]")));
		itemvalue.selectByVisibleText("Equals");
		Thread.sleep(6000);
		WebElement item = browser.findElement(By.xpath("//*[contains(@id,'Advan1:0:efqrp:value00::content')]"));
		waitUntilElementClickable("item", item, browser, timeout);
		WaituntilElementwritable("item", item, browser, Item);
		}
		catch(Exception e)
		{
		WebElement searchicon2 = browser.findElement(By.xpath("//*[contains(@id,'Advan1:0:efqrp::_afrDscl')]"));
		waitUntilElementClickable("searchicon2", searchicon2, browser, timeout);
		Thread.sleep(6000);
		Select itemvalue = new Select(browser.findElement(By.xpath("//*[contains(@id,'efqrp:operator0::content')]")));
		itemvalue.selectByVisibleText("Equals");
		Thread.sleep(6000);
		WebElement item = browser.findElement(By.xpath("//*[contains(@id,'Advan1:0:efqrp:value00::content')]"));
		waitUntilElementClickable("item", item, browser, timeout);
		WaituntilElementwritable("item", item, browser, Item);
		}
		WebElement searchicon = browser.findElement(By.xpath("//*[contains(@id,'Advan1:0:efqrp::search')]"));
		waitUntilElementClickable("searchicon", searchicon, browser, timeout);
		WebElement itemtable = browser.findElement(By.xpath("//*[contains(@id,'rstab:_ATp:table1::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("itemtable", itemtable, browser, timeout);
		WebElement ok = browser.findElement(By.xpath("//*[contains(@id,'itemNumberId2:cb1')]"));
		waitUntilElementClickable("qnty1", ok, browser, timeout);
		Thread.sleep(12000);
		WebElement addbutton = browser.findElement(By.xpath("//span[text()='Add']"));
		waitUntilElementClickable("addbutton", addbutton, browser, timeout);
		Thread.sleep(8000);
		List<WebElement> tablerow = browser.findElements(By.xpath("//*[contains(@id,'pc1:t1::db')]/table/tbody/tr"));
		int rowvalue = tablerow.size();
		System.out.println("The size of table is :" + rowvalue);
		WebElement itemvalue = browser.findElement(By.xpath("//*[text()='"+Item+"']"));
		waitUntilElementClickable("itemvalue", itemvalue, browser, timeout);
		Thread.sleep(3000);
		// WebElement qnty = browser.findElement(By.xpath("//*[contains(@id,'pc1:t1::db')]/table/tbody/tr["+m+"]/td[2]/div/table/tbody/tr/td[7]"));
		WebElement qnty = browser.findElement(By.xpath("//*[text()='"+Item+"']/../../../../../../../../../../../../../../..//input[contains(@id,'lineQuantity::content')]"));
		waitUntilElementClickable("qnty", qnty, browser, timeout);
		qnty.sendKeys("");
		qnty.sendKeys(Keys.DELETE);
		Thread.sleep(2000);
		qnty.sendKeys(Quantity);
		Thread.sleep(3000);
		WebElement el = browser.findElement(By.xpath("(//*[text()='"+Item+"']/../../../../../../../../../../../../../../..//img[contains(@id,'cil1::icon')])[1]"));
		JavascriptExecutor js = (JavascriptExecutor)browser;
		js.executeScript("arguments[0].scrollIntoView()", el);
		waitUntilElementClickable("el", el, browser, timeout);
		Thread.sleep(6000);
		if(!Adjustment_Type.equals("NA"))
		{
		Thread.sleep(2000);
		Select type = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc2::content')]")));
		type.selectByVisibleText(Adjustment_Type);
		}
		Thread.sleep(3000);
		if(!Unit_Selling_Price.equals("NA"))
		{
		WebElement sellingprice = browser.findElement(By.xpath("//*[contains(@id,'it2::content')]"));
		waitUntilElementClickable("sellingprice", sellingprice, browser, timeout);
		Thread.sleep(4000);
		sellingprice.sendKeys(Unit_Selling_Price);
		}
		if(!Reason.equals("NA"))
		{
		Select reason = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc1::content')]")));
		reason.selectByVisibleText(Reason);
		}
		Thread.sleep(4000);
		WebElement saveclosebutton = browser.findElement(By.xpath("(//*[text()='ave and Close'])[2]"));
		waitUntilElementClickable("saveclosebutton", saveclosebutton, browser, timeout);
		Thread.sleep(8000);
		WebElement actionbutton = browser.findElement(By.xpath("//*[text()='"+Item+"']/../../../../../../../../../../../../../../..//*[contains(@id,'gearIconCreate')]"));
		JavascriptExecutor js2 = (JavascriptExecutor)browser;
		js2.executeScript("arguments[0].scrollIntoView()", actionbutton);
		waitUntilElementClickable("actionbutton", actionbutton, browser, timeout);
		try {
		WebElement ma = browser.findElement(By.xpath("(//*[text()='Manage Attachments'])[2]"));
		waitUntilElementClickable("ma", ma, browser, timeout);
		}
		catch(Exception e)
		{
		WebElement ma = browser.findElement(By.xpath("(//*[text()='Manage Attachments'])[3]"));
		waitUntilElementClickable("ma", ma, browser, timeout);
		}
		if(!Manage_Attachements2.equals("NA"))
		{
		WebElement createicon = browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:create::icon')]"));
		waitUntilElementClickable("createicon", createicon, browser, timeout);
		WebElement createtable = browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("createtable", createtable, browser, timeout);
		Thread.sleep(2000);
		Select createtype = new Select(browser.findElement(By.xpath("//*[contains(@id,'dCode::content')]")));
		createtype.selectByVisibleText("File");
		browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable::db')]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]")).click();
		Thread.sleep(4000);
		Robot robo1 = new Robot();
		StringSelection str1 = new StringSelection("D:\\"+Manage_Attachements2+"");
		Thread.sleep(4000);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str1, null);
		robo1.keyPress(KeyEvent.VK_CONTROL);
		robo1.keyPress(KeyEvent.VK_V);
		robo1.keyRelease(KeyEvent.VK_CONTROL);
		robo1.keyRelease(KeyEvent.VK_V);
		robo1.keyPress(KeyEvent.VK_ENTER);
		robo1.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(3000);
		System.out.println("File is uploaded successfully");
		}
		Thread.sleep(6000);
		WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'manageAttachmentsDialog::ok')]"));
		waitUntilElementClickable("okbutton", okbutton, browser, timeout);
//		Thread.sleep(2000);
//		browser.navigate().refresh();
		Thread.sleep(8000);
		WebElement act = browser.findElement(By.xpath("//*[text()='"+Item+"']/../../../../../../../../../../../../../../..//*[contains(@id,'gearIconCreate')]"));
		JavascriptExecutor js4 = (JavascriptExecutor)browser;
		js4.executeScript("arguments[0].scrollIntoView()", act);
		waitUntilElementClickable("act", act, browser, timeout);
		try {
		Thread.sleep(3000);
		WebElement edi = browser.findElement(By.xpath("(//td[text()='Edit Additional Information'])[2]"));
		waitUntilElementClickable("edi", edi, browser, timeout);
		}
		catch(Exception e)
		{
		WebElement edi = browser.findElement(By.xpath("(//td[text()='Edit Additional Information'])[3]"));
		waitUntilElementClickable("edi", edi, browser, timeout);
		}
		if(!Repromise_Date.equals("NA"))
		{
		WebElement promise = browser.findElement(By.xpath("//*[contains(@id,'rePromiseDate::content')]"));
		waitUntilElementClickable("promise", promise, browser, timeout);
		promise.clear();
		WaituntilElementwritable("promise", promise, browser, Repromise_Date);
		}
		if(!Original_Schedule_Ship_Date.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'vantageScheduledShipDate::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'vantageScheduledShipDate::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'vantageScheduledShipDate::content')]")).sendKeys(Original_Schedule_Ship_Date);
		}
		if(!Customer_Catalouge_Cross_Refrerence.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::content')]")).sendKeys(Customer_Catalouge_Cross_Refrerence);
		}
		if(!Additional_Notes.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'additionalNotes::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'additionalNotes::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'additionalNotes::content')]")).sendKeys(Additional_Notes);
		}
		if(!MPN.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'mpn::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'mpn::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'mpn::content')]")).sendKeys(MPN);
		}
		if(!Internal_Item.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'internalItem::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'internalItem::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'internalItem::content')]")).sendKeys(Internal_Item);
		}
		if(!Cust_Src_Inspec.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'custSrcInspec::content')]")).click();
		}
		if(!Govt_Src_Inspec.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'govtSrcInspec::content')]")).click();
		}
		if(!FAI.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'fai::content')]")).click();
		}
		if(!Material_Certs.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'materialCerts::content')]")).click();
		}
		if(!Test_Reports.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'testReports::content')]")).click();
		}
		if(!Dimensional_Inspection.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'dimensionalInspection::content')]")).click();
		}
		if(!FAA_form.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'a81303FaaForm::content')]")).click();
		}
		Thread.sleep(2000);
		browser.findElement(By.linkText("Trading Partner Attributes")).click();
		Thread.sleep(2000);
		if(!SubIndustry_Segment.equals("NA"))
		{
				WebElement ackpopup3 = browser.findElement(By.xpath("//*[contains(@id,'subIndustrySegment_Display::lovIconId')]"));
				waitUntilElementClickable("ackpopup3", ackpopup3, browser, timeout);
				WebElement acksearch3 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch3", acksearch3, browser, timeout);
				WebElement ackvalue3 = browser.findElement(By.xpath("//*[contains(@id,'subIndustrySegment_Display::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("ackvalue3", ackvalue3, browser, timeout);
				ackvalue3.clear();
				WaituntilElementwritable("ackvalue3", ackvalue3, browser, SubIndustry_Segment);
				browser.findElement(By.xpath("//*[contains(@id,'subIndustrySegment_Display::_afrLovInternalQueryId::search')]")).click();
				WebElement acktable3 = browser.findElement(By.xpath("//*[contains(@id,'subIndustrySegment_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("acktable3", acktable3, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'subIndustrySegment_Display::lovDialogId::ok')]")).click();
		}
		if(!Program.equals("NA"))
		{
				WebElement ackpopup4 = browser.findElement(By.xpath("//*[contains(@id,'program_Display::lovIconId')]"));
				waitUntilElementClickable("ackpopup4", ackpopup4, browser, timeout);
				WebElement acksearch4 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch4", acksearch4, browser, timeout);
				WebElement ackvalue4 = browser.findElement(By.xpath("//*[contains(@id,'program_Display::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("ackvalue4", ackvalue4, browser, timeout);
				ackvalue4.clear();
				WaituntilElementwritable("ackvalue4", ackvalue4, browser, Program);
				browser.findElement(By.xpath("//*[contains(@id,'program_Display::_afrLovInternalQueryId::search')]")).click();
				WebElement acktable4 = browser.findElement(By.xpath("//*[contains(@id,'program_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("acktable4", acktable4, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'program_Display::lovDialogId::ok')]")).click();
		}
		if(!Platform.equals("NA"))
		{
				WebElement ackpopup5 = browser.findElement(By.xpath("//*[contains(@id,'platform_Display::lovIconId')]"));
				waitUntilElementClickable("ackpopup5", ackpopup5, browser, timeout);
				WebElement acksearch5 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch5", acksearch5, browser, timeout);
				WebElement ackvalue5 = browser.findElement(By.xpath("//*[contains(@id,'platform_Display::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("ackvalue5", ackvalue5, browser, timeout);
				ackvalue5.clear();
				WaituntilElementwritable("ackvalue5", ackvalue5, browser, Platform);
				browser.findElement(By.xpath("//*[contains(@id,'platform_Display::_afrLovInternalQueryId::search')]")).click();
				WebElement acktable5 = browser.findElement(By.xpath("//*[contains(@id,'platform_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("acktable5", acktable5, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'platform_Display::lovDialogId::ok')]")).click();
		}
		if(!TradingPartnerItem.equals("NA"))
		{
				WebElement ackpopup6 = browser.findElement(By.xpath("//*[contains(@id,'tradingPartnerItem_Display::lovIconId')]"));
				waitUntilElementClickable("ackpopup6", ackpopup6, browser, timeout);
				WebElement acksearch6 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch6", acksearch6, browser, timeout);
				WebElement ackvalue6 = browser.findElement(By.xpath("//*[contains(@id,'tradingPartnerItem_Display::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("ackvalue6", ackvalue6, browser, timeout);
				ackvalue6.clear();
				WaituntilElementwritable("ackvalue6", ackvalue6, browser, TradingPartnerItem);
				browser.findElement(By.xpath("//*[contains(@id,'tradingPartnerItem_Display::_afrLovInternalQueryId::search')]")).click();
				WebElement acktable6 = browser.findElement(By.xpath("//*[contains(@id,'0:_FOTsr1:1:AP1:r9:1:r2:0:dynam1:2:CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOTrading__Partner__Attributes:0:tradingPartnerItem_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[2]"));
				waitUntilElementClickable("acktable6", acktable6, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'tradingPartnerItem_Display::lovDialogId::ok')]")).click();
		}
		
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'dEffAttr::ok')]")).click();
		Thread.sleep(6000);
		if(!PrimarySalesperson.equals("NA"))
		{
				WebElement ackpopup7 = browser.findElement(By.xpath("//*[contains(@id,'AP1:pc1:t1:0:pspnlinecombo::lovIconId')]"));
				waitUntilElementClickable("ackpopup7", ackpopup7, browser, timeout);
				WebElement acksearch7 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch7", acksearch7, browser, timeout);
				WebElement ackvalue7 = browser.findElement(By.xpath("//*[contains(@id,'AP1:pc1:t1:0:qryId6:value00::content')]"));
				waitUntilElementClickable("ackvalue7", ackvalue7, browser, timeout);
				ackvalue7.clear();
				WaituntilElementwritable("ackvalue7", ackvalue7, browser, PrimarySalesperson);
				browser.findElement(By.xpath("//*[contains(@id,'AP1:pc1:t1:0:qryId6::search')]")).click();
				WebElement acktable7 = browser.findElement(By.xpath("//*[contains(@id,':AP1:pc1:t1:0:resId8::db')]//table/tbody/tr/td[1]"));
				waitUntilElementClickable("acktable7", acktable7, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'AP1:pc1:t1:0:pspnlinecombo::lovDialogId::ok')]")).click();
		}
		Thread.sleep(4000);
		if(!PrimarySalesperson.equals("NA"))
		{
				WebElement ackpopup8 = browser.findElement(By.xpath("//*[contains(@id,'AP1:primarySalesPersonNameId::lovIconId')]"));
				waitUntilElementClickable("ackpopup8", ackpopup8, browser, timeout);
				WebElement acksearch8 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch8", acksearch8, browser, timeout);
				WebElement ackvalue8 = browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId5:value00::content')]"));
				waitUntilElementClickable("ackvalue8", ackvalue8, browser, timeout);
				ackvalue8.clear();
				WaituntilElementwritable("ackvalue8", ackvalue8, browser, PrimarySalesperson);
				browser.findElement(By.xpath("//*[contains(@id,'FOTsr1:1:AP1:qryId5::search')]")).click();
				WebElement acktable8 = browser.findElement(By.xpath("//*[contains(@id,'FOTsr1:1:AP1:resId7::db')]//table/tbody/tr/td[2]"));
				waitUntilElementClickable("acktable8", acktable8, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'primarySalesPersonNameId::lovDialogId::ok')]")).click();
		}

		Thread.sleep(4000);
		sheet.getRow(i).createCell(85).setCellValue("Pass");
		       Updatefile(f, wb);
		}
		
		   }
		else
		{
		System.out.println("File is already processed");
		}
		
		
		try
		{
		wb.close();
		}
		catch(Exception e)
		{
		
		}
		}
		
		
		
		
		
		//method declaration for Headerpart update
		
		private void Headerpartupdate() throws Exception
		{
		WebElement bu = browser.findElement(By.xpath("//select[contains(@id, 'soc3::content')]"));
		Select bussUnt = new Select(bu);
		bussUnt.selectByVisibleText(Business_Unit);
		Thread.sleep(4000);
		WebElement customer = browser.findElement(By.xpath("//*[contains(@id,'partyNameId::lovIconId')]"));
		waitUntilElementClickable("customer", customer, browser, timeout);
		WebElement searchicon = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("searchicon", searchicon, browser, timeout);
		WebElement name = browser.findElement(By.xpath("//*[contains(@id,'qryId1:value00::content')]"));
		waitUntilElementClickable("name", name, browser, timeout);
		WaituntilElementwritable("name", name, browser, Customer);
		WebElement searchbutton = browser.findElement(By.xpath("//*[contains(@id,'qryId1::search')]"));
		waitUntilElementClickable("searchbutton", searchbutton, browser, timeout);
		WebElement table = browser.findElement(By.xpath("//*[contains(@id,'resId1::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("table", table, browser, timeout);
		Thread.sleep(4000);
		WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'partyNameId::lovDialogId::ok')]"));
		waitUntilElementClickable("okbutton", okbutton, browser, timeout);
		Thread.sleep(12000);
		WebElement purchase = browser.findElement(By.xpath("//*[contains(@id,'it1::content')]"));
		waitUntilElementClickable("purchase", purchase, browser, timeout);
		WaituntilElementwritable("purchase", purchase, browser, PurchaseOrder);
		if(!OrderType.equals("NA"))
		{
		WebElement ortype = browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::lovIconId')]"));
		waitUntilElementClickable("ortype", ortype, browser, timeout);
		WebElement searchtype = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("searchtype", searchtype, browser, timeout);
		WebElement ordertype = browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::_afrLovInternalQueryId:value00::content')]"));
		waitUntilElementClickable("ordertype", ordertype, browser, timeout);
		WaituntilElementwritable("ordertype", ordertype, browser, OrderType);
		WebElement orsearch = browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::_afrLovInternalQueryId::search')]"));
		waitUntilElementClickable("orsearch", orsearch, browser, timeout);
		WebElement ortable = browser.findElement(By.xpath("//*[contains(@id,'orderType1Id_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("ortable", ortable, browser, timeout);
		WebElement orok = browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::lovDialogId::ok')]"));
		waitUntilElementClickable("orok", orok, browser, timeout);
		}
		Thread.sleep(7000);
		if(!Contact.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactNameId::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactNameId::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactNameId::content')]")).sendKeys(Contact);
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactNameId::content')]")).sendKeys(Keys.ENTER);
		Thread.sleep(2000);
		}
		if(!Contact_Method.equals("NA"))
		{
		Thread.sleep(8000);
		browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactPointId::content')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactPointId::dropdownPopup::dropDownContent::db')]/table/tbody/tr[2]/td[2]/span")).click();
		}
		if(!Ship_to_Address.equals("NA"))
		{
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'shipToAddress::lovIconId')]")).click();
		WebElement search = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("search", search, browser, timeout);
		WebElement address = browser.findElement(By.xpath("//*[contains(@id,'qryId2:value00::content')]"));
		waitUntilElementClickable("address", address, browser, timeout);
		WaituntilElementwritable("address", address, browser, Ship_to_Address);
		browser.findElement(By.xpath("//*[contains(@id,'qryId2::search')]")).click();
		WebElement tablerow = browser.findElement(By.xpath("//*[contains(@id,'resId4::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("tablerow", tablerow, browser, timeout);
		Thread.sleep(2000);
		browser.findElement(By.xpath("//*[contains(@id,'shipToAddress::lovDialogId::ok')]")).click();
		}
		if(!Bill_to_Customer.equals("NA"))
		{
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'billToPartyNameId::lovIconId')]")).click();
		WebElement bcsearch = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("bcsearch", bcsearch, browser, timeout);
		WebElement bcname = browser.findElement(By.xpath("//*[contains(@id,'btpQry:value00::content')]"));
		waitUntilElementClickable("bcname", bcname, browser, timeout);
		bcname.clear();
		WaituntilElementwritable("bcname", bcname, browser, Bill_to_Customer);
		browser.findElement(By.xpath("//*[contains(@id,'btpQry::search')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'btpTbl::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'billToPartyNameId::lovDialogId::ok')]")).click();
		}
		Thread.sleep(12000);
		browser.findElement(By.xpath("//*[contains(@id,'sdi3::icon')]")).click();
		if(!Bill_to_Address.equals("NA"))
		{
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'billToLocation::lovIconId')]")).click();
		WebElement basearch = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("basearch", basearch, browser, timeout);
		WebElement baaddress = browser.findElement(By.xpath("//*[contains(@id,'qryId2:value00::content')]"));
		waitUntilElementClickable("baaddress", baaddress, browser, timeout);
		WaituntilElementwritable("baaddress", baaddress, browser, Bill_to_Address);
		browser.findElement(By.xpath("//*[contains(@id,'qryId2::search')]")).click();
		WebElement tableba = browser.findElement(By.xpath("//*[contains(@id,'resId2::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("tableba", tableba, browser, timeout);
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'billToLocation::lovDialogId::ok')]")).click();
		}
		if(!Bill_To_Contact.equals("NA"))
		{
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'billToContact::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'billToContact::content')]")).sendKeys(Bill_To_Contact);
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'billToContact::content')]")).sendKeys(Keys.ENTER);
		}
		if(!Payment_Terms.equals("NA"))
		{
		Thread.sleep(4000);
		Select pt = new Select(browser.findElement(By.xpath("//*[contains(@id,'paymentTermId::content')]")));
		pt.selectByVisibleText(Payment_Terms);
		}
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'sdi1::icon')]")).click();
		Thread.sleep(4000);
		browser.findElement(By.xpath("(//a[text()='Actions'])[1]")).click();
		Thread.sleep(2000);
		browser.findElement(By.xpath("//td[text()='Manage Attachments']")).click();
		WebElement icon = browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:create::icon')]"));
		waitUntilElementClickable("icon", icon, browser, timeout);
		WebElement table1 = browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("table1", table1, browser, timeout);
		Thread.sleep(3000);
		Select file = new Select(browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable:0:dCode::content')]")));
		file.selectByVisibleText("File");
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable::db')]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]")).click();
		Thread.sleep(4000);
		Robot robo = new Robot();
		StringSelection str = new StringSelection("D:\\"+Manage_Attachments+"");
		Thread.sleep(4000);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
		robo.keyPress(KeyEvent.VK_CONTROL);
		robo.keyPress(KeyEvent.VK_V);
		robo.keyRelease(KeyEvent.VK_CONTROL);
		robo.keyRelease(KeyEvent.VK_V);
		robo.keyPress(KeyEvent.VK_ENTER);
		robo.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(3000);
		System.out.println("File is uploaded successfully");
		Thread.sleep(6000);
		if(!Manage_Attachments1.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:create::icon')]")).click();
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable::db')]/table/tbody/tr[1]/td[1]")).click();
		Thread.sleep(6000);
		Select file1 = new Select(browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable:1:dCode::content')]")));
		file1.selectByVisibleText("File");
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'applicationsTable:_ATp:attachmentTable::db')]/table/tbody/tr[1]/td[2]/div/table/tbody/tr/td[2]")).click();
		Thread.sleep(4000);
		Robot robo1 = new Robot();
		StringSelection str1 = new StringSelection("D:\\"+Manage_Attachments1+"");
		Thread.sleep(4000);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str1, null);
		robo1.keyPress(KeyEvent.VK_CONTROL);
		robo1.keyPress(KeyEvent.VK_V);
		robo1.keyRelease(KeyEvent.VK_CONTROL);
		robo1.keyRelease(KeyEvent.VK_V);
		robo1.keyPress(KeyEvent.VK_ENTER);
		robo1.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(3000);
		System.out.println("File is uploaded successfully");
		}
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'manageHeaderAttachmentsDialog::ok')]")).click();
		Thread.sleep(6000);
		browser.findElement(By.xpath("(//a[text()='Actions'])[1]")).click();
		WebElement edit = browser.findElement(By.xpath("//td[text()='Edit Additional Information']"));
		waitUntilElementClickable("edit", edit, browser, timeout);
		if(!EndCustomer_Name.equals("NA"))
		{
		WebElement dropdown = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOAdditional__Header__Information:0:endCustomerName_Display::lovIconId')]"));
		waitUntilElementClickable("dropdown", dropdown, browser, timeout);
		WebElement ddsearch = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("ddsearch", ddsearch, browser, timeout);
		WebElement value = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOAdditional__Header__Information:0:endCustomerName_Display::_afrLovInternalQueryId:value00::content')]"));
		waitUntilElementClickable("value", value, browser, timeout);
		value.clear();
		WaituntilElementwritable("value", value, browser, EndCustomer_Name);
		browser.findElement(By.xpath("//button[text()='Search']")).click();
		WebElement entable = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOAdditional__Header__Information:0:endCustomerName_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("entable", entable, browser, timeout);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOAdditional__Header__Information:0:endCustomerName_Display::lovDialogId::ok')]")).click();
		}
		if(!Sales_Order_Acknowledgement_required.equals("NA"))
		{
		Thread.sleep(3000);
		WebElement ackpopup = browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::lovIconId')]"));
		waitUntilElementClickable("ackpopup", ackpopup, browser, timeout);
		WebElement acksearch = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable("acksearch", acksearch, browser, timeout);
		WebElement ackvalue = browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::_afrLovInternalQueryId:value00::content')]"));
		waitUntilElementClickable("ackvalue", ackvalue, browser, timeout);
		ackvalue.clear();
		WaituntilElementwritable("ackvalue", ackvalue, browser, Sales_Order_Acknowledgement_required);
		browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::_afrLovInternalQueryId::search')]")).click();
		WebElement acktable = browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
		waitUntilElementClickable("acktable", acktable, browser, timeout);
		browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::lovDialogId::ok')]")).click();
		Thread.sleep(6000);
		}
		if(!Header_Notes.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'headerNotes::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'headerNotes::content')]")).sendKeys(Header_Notes);
		}
		if(!Quality_Rating.equals("NA"))
		{
		WebElement ele = browser.findElement(By.xpath("//*[contains(@id,'qualityRating::content')]"));
		JavascriptExecutor js = (JavascriptExecutor)browser;
		js.executeScript("arguments[0].scrollIntoView()", ele);
		ele.click();
		ele.sendKeys(Quality_Rating);
		}
		if(!BDE.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'bde_Display::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'bde_Display::content')]")).sendKeys(BDE);
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'bde_Display::content')]")).sendKeys(Keys.ENTER);
		}
		if(!DPAS_Agency.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'dpasAgency::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'dpasAgency::content')]")).sendKeys(DPAS_Agency);
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'dpasAgency::content')]")).sendKeys(Keys.ENTER);
		}
		if(!DPAS_Program_ID.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'dpasProgramId::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'dpasProgramId::content')]")).sendKeys(DPAS_Program_ID);
		}
		if(!DPAS_Rating.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'dpasRating::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'dpasRating::content')]")).sendKeys(DPAS_Rating);
		}
		if(!Government_Contract.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'governmentContract::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'governmentContract::content')]")).sendKeys(Government_Contract);
		}
		if(!ITAR_Restricted.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'itarRestricted::content')]")).click();
		}
		if(!Export_License_Reqd.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'exportLicenseReqd_Display::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'exportLicenseReqd_Display::content')]")).sendKeys(Export_License_Reqd);
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'exportLicenseReqd_Display::content')]")).sendKeys(Keys.ENTER);
		}
		if(!FARS.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVORegulatory:0:fars::content')]")).click();
		}
		if(!DFARS.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVORegulatory:0:dfars::content')]")).click();
		}
		if(!Group.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'groupEngProd::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'groupEngProd::content')]")).sendKeys(Group);
		}
		if(!Region.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'region::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'region::content')]")).sendKeys(Region);
		}
		if(!Region2.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'region2::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'region2::content')]")).sendKeys(Region2);
		}
		Thread.sleep(3000);
		if(!SubEnd_Customer.equals("NA"))
		{
				WebElement ackpopup1 = browser.findElement(By.xpath("//*[contains(@id,'subEndCustomer_Display::lovIconId')]"));
				waitUntilElementClickable("ackpopup1", ackpopup1, browser, timeout);
				WebElement acksearch1 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch1", acksearch1, browser, timeout);
				WebElement ackvalue1 = browser.findElement(By.xpath("//*[contains(@id,'subEndCustomer_Display::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("ackvalue1", ackvalue1, browser, timeout);
				ackvalue1.clear();
				WaituntilElementwritable("ackvalue1", ackvalue1, browser, SubEnd_Customer);
				browser.findElement(By.xpath("//*[contains(@id,'subEndCustomer_Display::_afrLovInternalQueryId::search')]")).click();
				WebElement acktable1 = browser.findElement(By.xpath("//*[contains(@id,'subEndCustomer_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("acktable1", acktable1, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'subEndCustomer_Display::lovDialogId::ok')]")).click();
		}
		if(!PricingSubEnd_Customer.equals("NA"))
		{
				WebElement ackpopup2 = browser.findElement(By.xpath("//*[contains(@id,'PricingSubEndCustomer_Display::lovIconId')]"));
				waitUntilElementClickable("ackpopup2", ackpopup2, browser, timeout);
				WebElement acksearch2 = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("acksearch2", acksearch2, browser, timeout);
				WebElement ackvalue2 = browser.findElement(By.xpath("//*[contains(@id,'PricingSubEndCustomer_Display::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("ackvalue2", ackvalue2, browser, timeout);
				ackvalue2.clear();
				WaituntilElementwritable("ackvalue2", ackvalue2, browser, SubEnd_Customer);
				browser.findElement(By.xpath("//*[contains(@id,'PricingSubEndCustomer_Display::_afrLovInternalQueryId::search')]")).click();
				WebElement acktable2 = browser.findElement(By.xpath("//*[contains(@id,'PricingSubEndCustomer_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("acktable2", acktable2, browser, timeout);
				browser.findElement(By.xpath("//*[contains(@id,'PricingSubEndCustomer_Display::lovDialogId::ok')]")).click();
		}
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'dEffAttr::ok')]")).click();
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'sdi2::icon')]")).click();
		if(!Ship_To_Contact.equals("NA"))
		{
		WebElement shiptocontact = browser.findElement(By.xpath("//*[contains(@id,'shipToContactNameId::content')]"));
		waitUntilElementClickable("shiptocontact", shiptocontact, browser, timeout);
		shiptocontact.clear();
		WaituntilElementwritable("shiptocontact", shiptocontact, browser, Ship_To_Contact);
		Thread.sleep(5000);
		shiptocontact.sendKeys(Keys.ENTER);
		}
		if(!Ship_to_Contact_Method.equals("NA"))
		{
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'shipToContactPointId::content')]")).click();
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'shipToContactPointId::dropdownPopup::dropDownContent::db')]/table/tbody/tr[2]/td[2]")).click();
		}
		if(!ShippingMethod.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'shipMethodId::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'shipMethodId::content')]")).sendKeys(ShippingMethod);
		}
		if(!Requested_Date_Header.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).clear();
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
		 Calendar cal = Calendar.getInstance();
		browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).sendKeys(dateFormat.format(cal.getTime())+" 09:15 PM");
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).sendKeys(Keys.ENTER);
		}
		if(!Request_Type_Header.equals("NA"))
		{
		Thread.sleep(3000);
		Select rt = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc1::content')]")));
		rt.selectByVisibleText(Request_Type_Header);
		}
		Thread.sleep(6000);
		browser.findElement(By.linkText("Shipping")).click();
		if(!FOB_Header.equals("NA"))
		{
		Thread.sleep(6000);
		Select fob = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc3::content')]")));
		fob.selectByVisibleText(FOB_Header);
		}
		if(!FrieghtTerms_Header.equals("NA"))
		{
		Thread.sleep(3000);
		Select ft = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc4::content')]")));
		ft.selectByVisibleText(FrieghtTerms_Header);
		}
		if(!Shipment_Priority_Header.equals("NA"))
		{
		Select sp = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc5::content')]")));
		sp.selectByVisibleText(Shipment_Priority_Header);
		}
		if(!Shipping_Instructions_Header.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:it2::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:it2::content')]")).sendKeys(Shipping_Instructions_Header);
		}
		if(!Packing_Instructions_header.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:it3::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:it3::content')]")).sendKeys(Packing_Instructions_header);
		}
		Thread.sleep(6000);
		browser.findElement(By.linkText("Supply")).click();
		if(!Warehouse_Header.equals("NA"))
		{
		WebElement warehouse = browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]"));
		waitUntilElementClickable("warehouse", warehouse, browser, timeout);
		WaituntilElementwritable("warehouse", warehouse, browser, Warehouse_Header);
		}
		if(!Supplier_Header.equals("NA"))
		{
		browser.findElement(By.xpath("//*[contains(@id,'supplierNameId::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'supplierNameId::content')]")).sendKeys(Supplier_Header);
		}
		if(!Demand_Class_Header.equals("NA"))
		{
		Select demand = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc6::content')]")));
		demand.selectByVisibleText(Demand_Class_Header);
		}
		if(!Allow_partials.equals("NA"))
		{
		Select ap = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc7::content')]")));
		ap.selectByVisibleText(Allow_partials);
		}
		Thread.sleep(4000);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:sdi1::icon')]")).click();
		}
		
		
		
//		public void WebDriverwaitelement(WebElement element)
//		{
//		WebDriverWait wait = new WebDriverWait(browser,350);
//		wait.until(ExpectedConditions.visibilityOf(element));
//		}
		
		
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
		
		public static Integer getNumericValue(String str) {
		String str1[] = str.split("\\s");
		for (String s : str1) {
		boolean isNumeric = s.trim().chars().allMatch(Character::isDigit);
		if (isNumeric) {
		return Integer.parseInt(s);
		}
		}
		return 0;
		}
		
		
		
		public static void waitUntilElementClickable(String locatorName, final WebElement elementToWaitFor,
				WebDriver browser, int timeout) {
//			System.out.println("<<<<<< "+locatorName+">>>>>>>>");
			wait = new WebDriverWait(browser, timeout);
			wait.until(new Function<WebDriver, Boolean>() {
				int j;

				public Boolean apply(WebDriver browser) {
					j++;
					if (elementToWaitFor.isEnabled()) {
						try {
							elementToWaitFor.click();

						} catch (Exception e) {
							return false;

						}

					}
					return true;

				}
			});

		}
		
		public static void WaituntilElementwritable(String locatorName, final WebElement elementToWaitFor,
				WebDriver browser, String value) {
//			System.out.println("<<<<<< "+locatorName+" >>>>>>>>");

			wait = new WebDriverWait(browser, timeout);
			wait.until(new Function<WebDriver, Boolean>() {
				int j;

				public Boolean apply(WebDriver browser) {
					j++;
					if (elementToWaitFor.isEnabled()) {
						try {
							elementToWaitFor.sendKeys(value);

						} catch (Exception e) {
							return false;

						}

					}
					return true;

				}
			});

		}
		
		
		
		
		
		@AfterTest()
		public void Close_Browser()
		{
		// browser.close();
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