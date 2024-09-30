package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

public class Boyd_Order_Creation {
	
	FileInputStream fis;
	FileOutputStream fos;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	File srcFile;
	   WebDriver browser;
	   String result;
	   int i;
	   int j;
	  String orderNumber;
	  public String Header_Status;
	  public String Contact;
	  public String Method;
	  public int Purchase_orders;
	  public String Sales_Channel;
	  public String Sales_Order_Acknowledgement_Required;
	  public String DPAS_Agency;
	  public String DPAS_Program_ID;
	  public String ITAR_Restricted;
	  public String FARS;
	  public String DFARS;
	  public String Group;
	  public String Region;
	  public String Ship_to_Contact;
	  public String Ship_to_Contact_Method;
	  public String Request_Type;
	  public String Requested_Date;
	  public String Shipping_Method;
	  public String Ship_Lines_Together;
	  public String FOB;
	  public String Freight_Terms;
	  public String Allow_Partial_Shipments_of_Lines;
	  public String Shipment_Priority;
	  public String Shipping_Instructions;
	  public String Packing_Instructions;
	  public String Warehouse;
	  public String Allow_Item_Substitution;
	  public String Bill_to_Address;
	  public String Bill_to_Contact;
	  public String Bill_to_Contact_Method;
	  public String Payment_Terms;
	  public String Payment_Method;
	  public String Credit_Approved_in_Source_System;
	  public static String item;
	  public static String lineNumber;
	   public static String itemQnty;
	   public static String unitSellingPrice;
	   public static String type;
	   public static String pricingAdj;
	   public static String reason;
	   public static String lineAmt;
	   public static String scheduleShipDate;
	   public static String reqShipDate;
	   public static String reqType;
	   public static String fOB;
	   public static String freightTerms;
	   public static String shipPriority;
	   public static String shipInstructions;
	   public static String shipMethod;
	   public static String packInstructions;
	   public static String shipTOContact;
	   public static String shipTOCustomer;
	   public static String warehouse;
	   public static String purchseOrder;
	   public static String purchseOrderLine;
	   public static String recTransactions;
	   public static String lineEFF;
	   public static String rePromiseDate;
	   public static String origScheduleShipDate;
	   public static String catalogCrossRef;
	   public static String mPNOverride;
	   public static String additionalNote;
	   public static String custSrcInspec;
	   public static String govtSrcInspec;
	   public static String FAI;
	   public static String materialCert;
	   public static String testReport;
	   public static String dimInspec;
	   public static String FAA_Form;
	   JavascriptExecutor js = (JavascriptExecutor) browser;
	   public static boolean flag;
	   @BeforeTest
	   public void Login_Page() throws Exception
	   {
	   WebDriverManager.chromedriver().setup();
	   ChromeOptions options = new ChromeOptions();
	   options.setPageLoadStrategy(PageLoadStrategy.NONE);
	   browser = new ChromeDriver(options);
	   browser.manage().window().maximize();
	   browser.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
	   browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
//	    browser.get("https://elme-dev2.fa.us8.oraclecloud.com/");
	   browser.get("https://elme.fa.us8.oraclecloud.com/");
	   browser.findElement(By.id("userid")).click();
	   browser.findElement(By.id("userid")).sendKeys("forsys.user");
	   browser.findElement(By.id("password")).click();
	   browser.findElement(By.id("password")).sendKeys("Boyd2@2!");
	   browser.findElement(By.id("btnActive")).click();
	   Thread.sleep(5000);
	   WebDriverWait wait = new WebDriverWait(browser,350);
	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("pt1:_UIShome::icon")));
	   browser.findElement(By.id("pt1:_UIShome::icon")).click();
	Thread.sleep(7000);
	browser.findElement(By.id("groupNode_order_management")).click();
	browser.findElement(By.id("itemNode_order_management_order_management")).click();

	}
	   
	   @Test
	   public void updateOrder() throws Exception
	   {
	    try {
	    JavascriptExecutor js = (JavascriptExecutor) browser;  
	srcFile = new File(System.getProperty("user.dir")+"\\Excel\\Boyd_Order_Creation.xlsx");
	fis = new FileInputStream(srcFile);
	wb = new XSSFWorkbook(fis);
	sheet = wb.getSheetAt(0);
	int totalRows = sheet.getPhysicalNumberOfRows();
	System.out.println("Total number of Excel rows are :" +totalRows);
	for(i=1; i<=totalRows; )
	{
	if(sheet.getRow(i) == null) {
	try {
	JavascriptExecutor jse = (JavascriptExecutor) browser;
	      jse.executeScript("window.scrollBy(0,-350)", "");
	Thread.sleep(4000);
	JavascriptExecutor js1 = (JavascriptExecutor) browser;
	WebElement ele = browser.findElement(By.xpath("//a[contains(@id, 'APRS1:save::popEl')]"));
	js1.executeScript("arguments[0].scrollIntoView();",ele );
	ele.click();
	Thread.sleep(3000);
	browser.findElement(By.xpath("//*[contains(@id, 'APRS1:cmi1')]")).click();
	System.out.println("Order Saved");
	Thread.sleep(4000);
	browser.findElement(By.xpath("//div[contains(@id, 'APVIEW1:SPb')]")).click();
	Thread.sleep(4000);
	browser.findElement(By.xpath("//div[contains(@id, 'AP1:SPb')]")).click();
	i++;
	continue;
	}
	catch(Exception e)
	{
	i++;
	continue;
	}

	}

	if(sheet.getRow(i) != null && isRowEmpty(sheet.getRow(i))) {
	try {
	JavascriptExecutor jse = (JavascriptExecutor) browser;
	      jse.executeScript("window.scrollBy(0,-350)", "");
	Thread.sleep(4000);
	JavascriptExecutor js1 = (JavascriptExecutor) browser;
	WebElement ele = browser.findElement(By.xpath("//a[contains(@id, 'APRS1:save::popEl')]"));
	js1.executeScript("arguments[0].scrollIntoView();",ele );
	ele.click();
	Thread.sleep(3000);
	browser.findElement(By.xpath("//*[contains(@id, 'APRS1:cmi1')]")).click();
	System.out.println("Order Saved");
	Thread.sleep(4000);
	browser.findElement(By.xpath("//div[contains(@id, 'APVIEW1:SPb')]")).click();
	Thread.sleep(4000);
	browser.findElement(By.xpath("//div[contains(@id, 'AP1:SPb')]")).click();
	i++;
	continue;
	}
	catch(Exception e)
	{
	i++;
	continue;
	}
	}
	orderNumber = sheet.getRow(i).getCell(0).getStringCellValue().trim();
	Header_Status = sheet.getRow(i).getCell(1).getStringCellValue().trim();
	Contact = sheet.getRow(i).getCell(2).getStringCellValue().trim();
	Method = sheet.getRow(i).getCell(3).getStringCellValue().trim();
	// Purchase_orders = (int)sheet.getRow(i).getCell(4).getNumericCellValue();
	String purchase = sheet.getRow(i).getCell(4).getStringCellValue().trim();
	Sales_Channel = sheet.getRow(i).getCell(5).getStringCellValue();
	Sales_Order_Acknowledgement_Required = sheet.getRow(i).getCell(6).getStringCellValue().trim();
	DPAS_Agency = sheet.getRow(i).getCell(7).getStringCellValue().trim();
	DPAS_Program_ID = sheet.getRow(i).getCell(8).getStringCellValue().trim();
	ITAR_Restricted = sheet.getRow(i).getCell(9).getStringCellValue().trim();
	FARS = sheet.getRow(i).getCell(10).getStringCellValue().trim();
	DFARS = sheet.getRow(i).getCell(11).getStringCellValue().trim();
	Group = sheet.getRow(i).getCell(12).getStringCellValue().trim();
	Region = sheet.getRow(i).getCell(13).getStringCellValue().trim();
	Ship_to_Contact = sheet.getRow(i).getCell(14).getStringCellValue().trim();
	Ship_to_Contact_Method = sheet.getRow(i).getCell(15).getStringCellValue().trim();
	Request_Type = sheet.getRow(i).getCell(16).getStringCellValue().trim();
	Requested_Date = sheet.getRow(i).getCell(17).getStringCellValue().trim();
	Shipping_Method = sheet.getRow(i).getCell(18).getStringCellValue().trim();
	Ship_Lines_Together = sheet.getRow(i).getCell(19).getStringCellValue().trim();
	FOB = sheet.getRow(i).getCell(20).getStringCellValue().trim();
	Freight_Terms = sheet.getRow(i).getCell(21).getStringCellValue().trim();
	Allow_Partial_Shipments_of_Lines = sheet.getRow(i).getCell(22).getStringCellValue().trim();
	Shipment_Priority = sheet.getRow(i).getCell(23).getStringCellValue().trim();
	Shipping_Instructions = sheet.getRow(i).getCell(24).getStringCellValue().trim();
	Packing_Instructions = sheet.getRow(i).getCell(25).getStringCellValue().trim();
	Warehouse = sheet.getRow(i).getCell(26).getStringCellValue().trim();
	Allow_Item_Substitution = sheet.getRow(i).getCell(27).getStringCellValue().trim();
	Bill_to_Address = sheet.getRow(i).getCell(28).getStringCellValue().trim();
	Bill_to_Contact = sheet.getRow(i).getCell(29).getStringCellValue().trim();
	Bill_to_Contact_Method = sheet.getRow(i).getCell(30).getStringCellValue().trim();
	Payment_Terms = sheet.getRow(i).getCell(31).getStringCellValue().trim();
	Payment_Method = sheet.getRow(i).getCell(32).getStringCellValue().trim();
	Credit_Approved_in_Source_System = sheet.getRow(i).getCell(33).getStringCellValue().trim();

	item = sheet.getRow(i).getCell(35).getStringCellValue().trim();



	if(!orderNumber.equals("NA") && !orderNumber.isEmpty())
	{
	flag=true;
	Thread.sleep(3000);
	search();
	Thread.sleep(7000);
	 browser.findElement(By.linkText(orderNumber)).click();
	 Thread.sleep(5000);
	 browser.findElement(By.xpath("//*[contains(@id, 'APVIEW1:m1')]/div/table/tbody/tr/td[3]/div")).click();
	 Thread.sleep(3000);
	 browser.findElement(By.xpath("//tr[contains(@id, 'APVIEW1:createRevision')]")).click();
	 Thread.sleep(8000);
	 
	 //**Header Part Updates**//
	 
	 if(!Contact.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:soldToPartyContactNameId::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:soldToPartyContactNameId::content')]")).sendKeys(Contact);
	 Thread.sleep(4000);
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:soldToPartyContactNameId::content')]")).sendKeys(Keys.ENTER);
	 }
	 if(!Method.equals("NA"))
	 {
	 Thread.sleep(4000);
	 browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactPointId::lovIconId')]")).click();
	 Thread.sleep(2000);
	 browser.findElement(By.xpath("//*[contains(@id,'soldToPartyContactPointId::dropdownPopup::dropDownContent::db')]/table/tbody/tr[2]/td[2]/span")).click();
	 Thread.sleep(4000);
	 }
	 if(!purchase.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:it1::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:it1::content')]")).sendKeys(purchase);
	 }
	 Thread.sleep(4000);
	 if(!Sales_Channel.equals("NA"))
	 {
	 Select sc1 = new Select(browser.findElement(By.xpath("//*[contains(@id,'APRS1:soc5::content')]")));
	 sc1.selectByVisibleText(Sales_Channel);
	 }
	 Thread.sleep(4000);
	 browser.findElement(By.xpath("//a[text()='Actions']")).click();
	 Thread.sleep(2000);
	 browser.findElement(By.xpath("//td[text()='Edit Additional Information']")).click();
	 Thread.sleep(4000);
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::lovIconId')]")).click();
	 Thread.sleep(2000);
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::dropdownPopup::popupsearch')]")).click();
	 Thread.sleep(2000);
	 if(!Sales_Order_Acknowledgement_Required.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::_afrLovInternalQueryId:value00::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::_afrLovInternalQueryId:value00::content')]")).clear();
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::_afrLovInternalQueryId:value00::content')]")).sendKeys(Sales_Order_Acknowledgement_Required);
	 }
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::_afrLovInternalQueryId::search')]")).click();
	 Thread.sleep(3000);
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'salesOrderAcknowledgementRequi::lovDialogId::ok')]")).click();
	 Thread.sleep(4000);
	 if(!DPAS_Agency.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'dpasAgency::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'dpasAgency::content')]")).sendKeys(DPAS_Agency);
	 }
	 Thread.sleep(3000);
	 if(!DPAS_Program_ID.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'dpasProgramId::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'dpasProgramId::content')]")).sendKeys(DPAS_Program_ID);
	 }
	 Thread.sleep(3000);
	 if(!ITAR_Restricted.equals("NA"))
	{
	browser.findElement(By.xpath("//*[contains(@id,'itarRestricted::Label1')]")).click();
	}
	Thread.sleep(2000);
	if(!FARS.equals("NA"))
	{
	browser.findElement(By.xpath("(//*[contains(@id,'fars::Label1')])[1]")).click();
	}
	Thread.sleep(2000);
	if(!DFARS.equals("NA"))
	{
	browser.findElement(By.xpath("//*[contains(@id,'dfars::Label1')]")).click();
	}
	 Thread.sleep(4000);
	 if(!Group.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'groupEngProd::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'groupEngProd::content')]")).sendKeys(Group);
	 }
	 Thread.sleep(4000);
	 if(!Region.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'region::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'region::content')]")).sendKeys(Region);
	 }
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:dEffAttr::ok')]")).click();
	 Thread.sleep(6000);
	 browser.findElement(By.xpath("//*[contains(@id,'APRS1:sdi2::icon')]")).click();
	 Thread.sleep(6000);
	 browser.findElement(By.linkText("General")).click();
	 Thread.sleep(4000);
	 if(!Ship_to_Contact.equals("NA"))
	 {
	 Thread.sleep(4000);
	 browser.findElement(By.xpath("//*[contains(@id,'shipToContactNameId::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'shipToContactNameId::content')]")).sendKeys(Ship_to_Contact);
	 Thread.sleep(3000);
	 browser.findElement(By.xpath("//*[contains(@id,'shipToContactNameId::content')]")).sendKeys(Keys.ENTER);
	 }
	 if(!Ship_to_Contact_Method.equals("NA"))
	 {
	 Thread.sleep(4000);
	 browser.findElement(By.xpath("//*[contains(@id,'shipToContactPointId::lovIconId')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'shipToContactPointId::dropdownPopup::dropDownContent::db')]/table/tbody/tr[2]/td[2]/span")).click();
	 Thread.sleep(4000);
	 }
	 if(!Request_Type.equals("NA"))
	 {
	 Thread.sleep(4000);
	 Select request = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc1::content')]")));
	 request.selectByVisibleText(Request_Type);
	 }
	 Thread.sleep(2000);
	 if(!Requested_Date.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).clear();
	 browser.findElement(By.xpath("//*[contains(@id,'id1::content')]")).sendKeys(Requested_Date);
	 }
	 Thread.sleep(3000);
	 if(!Shipping_Method.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'shipMethodId::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'shipMethodId::content')]")).sendKeys(Shipping_Method);
	 }
	 Thread.sleep(6000);
	 browser.findElement(By.linkText("Shipping")).click();
	 Thread.sleep(4000);
	 if(!FOB.equals("NA"))
	 {
	 Select fo = new Select(browser.findElement(By.xpath("(//*[contains(@id,'soc3::content')])[2]")));
	 fo.selectByVisibleText(FOB);
	 }
	 Thread.sleep(3000);
	 if(!Freight_Terms.equals("NA"))
	 {
	 Select fr = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc4::content')]")));
	 fr.selectByVisibleText(Freight_Terms);
	 }
	 Thread.sleep(3000);
	 if(!Allow_Partial_Shipments_of_Lines.equals("NA"))
	         {
	 Select apl = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc2::content')]")));
	 apl.selectByVisibleText(Allow_Partial_Shipments_of_Lines);
	         }
	 Thread.sleep(3000);
	 if(!Shipment_Priority.equals("NA"))
	 {
	 Select sp = new Select(browser.findElement(By.xpath("(//*[contains(@id,'soc5::content')])[2]")));
	 sp.selectByVisibleText(Shipment_Priority);
	 }
	 Thread.sleep(3000);
	 if(!Shipping_Instructions.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'it2::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'it2::content')]")).sendKeys(Shipping_Instructions);
	 }
	 Thread.sleep(3000);
	 if(!Packing_Instructions.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'it3::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'it3::content')]")).sendKeys(Packing_Instructions);
	 }
	 Thread.sleep(10000);
	 browser.findElement(By.linkText("Supply")).click();
	 Thread.sleep(4000);
	 if(!Warehouse.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]")).clear();
	 browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]")).sendKeys(Warehouse);
	 }
	 Thread.sleep(3000);
	 if(!Allow_Item_Substitution.equals("NA"))
	 {
	 Select ais = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc7::content')]")));
	 ais.selectByVisibleText(Allow_Item_Substitution);
	 }
	 Thread.sleep(6000);
	 browser.findElement(By.xpath("//*[contains(@id,'sdi3::icon')]")).click();
	 Thread.sleep(6000);
	 if(!Bill_to_Address.equals("NA"))
	 {
	 Select BA = new Select(browser.findElement(By.xpath("//*[contains(@id,'billToLocation::content')]")));
	 BA.selectByVisibleText(Bill_to_Address);
	 }
	 Thread.sleep(6000);
	 if(!Bill_to_Contact.equals("NA"))
	 {
	 browser.findElement(By.xpath("//*[contains(@id,'billToContact::content')]")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'billToContact::content')]")).sendKeys(Bill_to_Contact);
	 browser.findElement(By.xpath("//*[contains(@id,'billToContact::content')]")).sendKeys(Keys.ENTER);
	 }
	 if(!Bill_to_Contact_Method.equals("NA"))
	 {
	 Thread.sleep(6000);
	 browser.findElement(By.xpath("//*[contains(@id,'billToContactPointId::lovIconId')]")).click();
	 Thread.sleep(3000);
	 browser.findElement(By.xpath("//*[contains(@id,'billToContactPointId::dropdownPopup::dropDownContent::db')]/table/tbody/tr[2]/td[2]/span")).click();
	 }
	 if(!Payment_Terms.equals("NA"))
	 {
	 Thread.sleep(3000);
	 Select pt = new Select(browser.findElement(By.xpath("//*[contains(@id,'paymentTermId::content')]")));
	 pt.selectByVisibleText(Payment_Terms);
	 }
	}  
	 //**Lines Update Started**//
	 if(!item.equalsIgnoreCase("NA"))
	{
	 Thread.sleep(5000);
	browser.findElement(By.xpath("//img[contains(@id, 'APRS1:sdi1::icon')]")).click();
	try {
	WebElement ele1 = browser.findElement(By.xpath("//input[contains(@id, 'APRS1_afr_pc1_afr_t1_afr_c2::content')]"));
	js.executeScript("arguments[0].scrollIntoView();",ele1 );
	ele1.clear();
	ele1.sendKeys(item);
	ele1.sendKeys(Keys.ENTER);
	}
	catch(Exception e)
	{
	browser.findElement(By.xpath("//img[contains(@id, 'APRS1:pc1:_qbeTbr::icon')]")).click();
	WebElement ele1 = browser.findElement(By.xpath("//input[contains(@id, 'APRS1_afr_pc1_afr_t1_afr_c2::content')]"));
	js.executeScript("arguments[0].scrollIntoView();",ele1 );
	ele1.clear();
	ele1.sendKeys(item);
	ele1.sendKeys(Keys.ENTER);
	}
	Thread.sleep(5000);
	List<WebElement> table = browser.findElements(By.xpath("//*[contains(@id, 'APRS1:pc1:t1::db')]/table/tbody/tr"));
	int tablesize = table.size();
	System.out.println("tablesize="+tablesize);
	if(tablesize==0)
	{
	System.out.println("Item not found");
	XSSFCell cell1 = sheet.getRow(i).createCell(49);
	cell1.setCellValue("Item Number not available");
	fos = new FileOutputStream(srcFile);
	wb.write(fos);
	i++;
	}
	else {
	for(j=1; j<=tablesize; j++)
	{
	try {
	lineNumber = sheet.getRow(i).getCell(34).getStringCellValue().trim();
	item = sheet.getRow(i).getCell(35).getStringCellValue().trim();
	unitSellingPrice = sheet.getRow(i).getCell(36).getStringCellValue().trim();
	type = sheet.getRow(i).getCell(37).getStringCellValue().trim();
	pricingAdj = sheet.getRow(i).getCell(38).getStringCellValue().trim();
	reason = sheet.getRow(i).getCell(39).getStringCellValue().trim();
	lineAmt = sheet.getRow(i).getCell(40).getStringCellValue().trim();

	reqShipDate = sheet.getRow(i).getCell(41).getStringCellValue().trim();
	reqType = sheet.getRow(i).getCell(42).getStringCellValue().trim();
	fOB = sheet.getRow(i).getCell(43).getStringCellValue().trim();
	freightTerms = sheet.getRow(i).getCell(44).getStringCellValue().trim();
	shipPriority = sheet.getRow(i).getCell(45).getStringCellValue().trim();
	shipInstructions = sheet.getRow(i).getCell(46).getStringCellValue().trim();
	shipMethod = sheet.getRow(i).getCell(47).getStringCellValue().trim();
	packInstructions = sheet.getRow(i).getCell(48).getStringCellValue().trim();

	warehouse = sheet.getRow(i).getCell(49).getStringCellValue().trim();
	purchseOrder = sheet.getRow(i).getCell(50).getStringCellValue().trim();
	purchseOrderLine = sheet.getRow(i).getCell(51).getStringCellValue().trim();
	recTransactions = sheet.getRow(i).getCell(52).getStringCellValue().trim();
	rePromiseDate = sheet.getRow(i).getCell(53).getStringCellValue().trim();
	origScheduleShipDate = sheet.getRow(i).getCell(54).getStringCellValue().trim();
	catalogCrossRef = sheet.getRow(i).getCell(55).getStringCellValue().trim();
	mPNOverride = sheet.getRow(i).getCell(56).getStringCellValue().trim();
	additionalNote = sheet.getRow(i).getCell(57).getStringCellValue().trim();
	custSrcInspec = sheet.getRow(i).getCell(58).getStringCellValue().trim();
	govtSrcInspec = sheet.getRow(i).getCell(59).getStringCellValue().trim();
	FAI = sheet.getRow(i).getCell(60).getStringCellValue().trim();
	materialCert = sheet.getRow(i).getCell(61).getStringCellValue().trim();
	testReport = sheet.getRow(i).getCell(62).getStringCellValue().trim();
	dimInspec = sheet.getRow(i).getCell(63).getStringCellValue().trim();
	FAA_Form = sheet.getRow(i).getCell(64).getStringCellValue().trim();
	shipTOContact = sheet.getRow(i).getCell(65).getStringCellValue().trim();
	scheduleShipDate = sheet.getRow(i).getCell(66).getStringCellValue().trim();
	}
	catch(Exception e)
	{
	e.printStackTrace();
	System.out.println("Unable to get data");
	}
	Thread.sleep(5000);
	    js.executeScript("window.scrollBy(0, 2000)", "");
	   
	    WebElement LineNum = browser.findElement(By.xpath("//*[contains(@id, 'APRS1:pc1:t1::db')]/table/tbody/tr["+j+"]/td[1]"));
	    String lineNumb = LineNum.getText().trim();
	    System.out.println("lineNumb=="+lineNumb);
	    if(lineNumb.equalsIgnoreCase(lineNumber))
	    {
	    browser.findElement(By.xpath("//*[contains(@id, 'APRS1:pc1:t1::db')]/table/tbody/tr["+j+"]/td[1]")).click();
	   Thread.sleep(3000);
	   
	    if(!pricingAdj.equals("NA") && !pricingAdj.isEmpty())
	{
	    unitSellingPrice();
	}
	    Thread.sleep(6000);
	    int n=0;
	    int a = 0,b = 0,c = 0,d = 0,e = 0,f = 0,g = 0,h = 0,o = 0,p = 0,q = 0,r = 0;
	    if(!reqShipDate.equals("NA") || !reqType.equals("NA") || !fOB.equals("NA") || !freightTerms.equals("NA") || !shipPriority.equals("NA")
	    ||!shipInstructions.equals("NA") || !packInstructions.equals("NA") || !warehouse.equals("NA") || !purchseOrder.equals("NA")
	    || !purchseOrderLine.equals("NA")||!recTransactions.equals("NA") ||!shipMethod.equals("NA"))
	    {
	    browser.findElement(By.xpath("//span[text()='Update Lines']")).click();
	    if(!reqShipDate.equals("NA"))
	    {
	    Thread.sleep(5000);
	    browser.findElement(By.xpath("//*[text()='Requested Date']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    a=n++;
	    }
	    if(!reqType.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Request Type']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    b=n++;
	    }
	    if(!fOB.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='FOB']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    c=n++;
	    }
	    if(!freightTerms.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Freight Terms']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    d=n++;
	    }
	    if(!shipPriority.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Shipment Priority']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    e=n++;
	    }
	    if(!shipInstructions.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Shipping Instructions']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    f=n++;
	    }
	    if(!shipMethod.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Shipping Method']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    g=n++;
	    }
	    if(!packInstructions.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Packing Instructions']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    h=n++;
	    }
	    if(!warehouse.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Warehouse']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    o=n++;
	    }
	    if(!purchseOrder.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Purchase Order']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    p=n++;
	    }
	    if(!purchseOrderLine.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Purchase Order Line']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    q=n++;
	    }
	    if(!recTransactions.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[text()='Receivables Transaction']")).click();
	    browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	    Thread.sleep(4000);
	    r=n++;
	    }
	   
	    browser.findElement(By.xpath("//*[text()='ext']")).click();
	    Thread.sleep(5000);
	    if(!reqShipDate.equals("NA"))
	    {
	     WebElement reqdate = browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+a+":id2::content')]"));
	     reqdate.sendKeys(reqShipDate);
	    Thread.sleep(3000);
	    }
	    if(!reqType.equals("NA"))
	    {
	     Select reqtyp = new Select(browser.findElement(By.xpath("//select[contains(@id, 'SP2:t2:"+b+":soc1::content')]")));
	     reqtyp.selectByVisibleText(reqType);
	    Thread.sleep(3000);
	    }
	    if(!fOB.equals("NA"))
	    {
	     Select fob1 = new Select(browser.findElement(By.xpath("//select[contains(@id, 'SP2:t2:"+c+":soc1::content')]")));
	    fob1.selectByVisibleText(fOB);
	    Thread.sleep(3000);
	    }
	    if(!freightTerms.equals("NA"))
	    {
	     WebElement reqdate = browser.findElement(By.xpath("//select[contains(@id, 'SP2:t2:"+d+":soc1::content')]"));
	     reqdate.sendKeys(freightTerms);
	    Thread.sleep(3000);
	    }
	    if(!shipPriority.equals("NA"))
	    {
	     Select reqtyp = new Select(browser.findElement(By.xpath("//select[contains(@id, 'SP2:t2:"+e+":soc1::content')]")));
	     reqtyp.selectByVisibleText(shipPriority);
	    Thread.sleep(3000);
	    }
	    if(!shipInstructions.equals("NA"))
	    {
	    WebElement ins = browser.findElement(By.xpath("//*[contains(@id, 'SP2:t2:"+f+":it2::content')]"));
	    ins.sendKeys(shipInstructions);
	    Thread.sleep(3000);
	    }
	    if(!shipMethod.equals("NA"))
	    {
	    WebElement ins = browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+g+":integerValueId::content')]"));
	    ins.sendKeys(shipMethod);
	    Thread.sleep(3000);
	    }
	    if(!packInstructions.equals("NA"))
	    {
	    WebElement pins = browser.findElement(By.xpath("//*[contains(@id, 'SP2:t2:"+h+":it2::content')]"));
	    pins.sendKeys(packInstructions);
	    Thread.sleep(3000);
	    }
	    if(!warehouse.equals("NA"))
	    {
	    WebElement ware = browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+o+":integerValueId::content')]"));
	    ware.sendKeys(warehouse);
	    Thread.sleep(3000);
	    }
	    if(!purchseOrder.equals("NA"))
	    {
//	     WebElement purchase1 = browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:8:it2::content')]"));
	    browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+p+":it2::content')]")).click();
	    browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+p+":it2::content')]")).clear();
	    browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+p+":it2::content')]")).sendKeys(purchseOrder);
	    browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+p+":it2::content')]")).sendKeys(Keys.ENTER);
	    Thread.sleep(5000);
	    }
	    if(!purchseOrderLine.equals("NA"))
	    {
	    WebElement purOrder = browser.findElement(By.xpath("//input[contains(@id, 'SP2:t2:"+q+":it2::content')]"));
	    purOrder.sendKeys(purchseOrderLine);
	    Thread.sleep(3000);
	    }
	    if(!recTransactions.equals("NA"))
	    {
	    Select recTran = new Select(browser.findElement(By.xpath("//select[contains(@id, 'SP2:t2:"+r+":soc1::content')]")));
	    recTran.selectByVisibleText(recTransactions);
	    Thread.sleep(3000);
	    }
	    Thread.sleep(5000);
	    browser.findElement(By.xpath("//*[text()='ave and Close']")).click();
	    Thread.sleep(4000);
	    try
	    {
	    browser.findElement(By.xpath("//button[contains(@id, 'SP2:cb1')]")).click();
	    }
	    catch(Exception e1)
	    {
	   
	    }
	    }
	    Thread.sleep(5000);
	    browser.findElement(By.xpath("//*[contains(@id,'sdi1::icon')]")).click();
	    Thread.sleep(6000);
	    WebElement ele = browser.findElement(By.xpath("(//*[contains(@id,'gearIconRevise')])["+j+"]"));
	    JavascriptExecutor js1 = (JavascriptExecutor)browser;
	    js1.executeScript("arguments[0].scrollIntoView()", ele);
	    ele.click();
	    Thread.sleep(4000);
	    browser.findElement(By.xpath("(//*[text()='Edit Additional Information'])[2]")).click();
	    Thread.sleep(4000);
	    if(!rePromiseDate.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'rePromiseDate::content')]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'rePromiseDate::content')]")).sendKeys(rePromiseDate);
	    }
	    Thread.sleep(4000);
	    if(!origScheduleShipDate.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'vantageScheduledShipDate::content')]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'vantageScheduledShipDate::content')]")).sendKeys(origScheduleShipDate);
	    }
	    Thread.sleep(4000);
	    if(!catalogCrossRef.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::lovIconId')]")).click();
	    Thread.sleep(4000);
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::dropdownPopup::popupsearch')]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::_afrLovInternalQueryId:value00::content')]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::_afrLovInternalQueryId:value00::content')]")).clear();
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::_afrLovInternalQueryId:value00::content')]")).sendKeys(catalogCrossRef);
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::_afrLovInternalQueryId::search')]")).click();
	    Thread.sleep(4000);
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'customerCatalogCrossReference_Display::lovDialogId::ok')]")).click();
	    Thread.sleep(6000);

	    }
	    if(!mPNOverride.equals("NA"))
	    {
	    browser.findElement(By.xpath("//input[contains(@id, 'Attributes:0:mpn::content')]")).click();
	    browser.findElement(By.xpath("//input[contains(@id, 'Attributes:0:mpn::content')]")).sendKeys(mPNOverride);
	    Thread.sleep(4000);
	    }
	    if(!additionalNote.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'additionalNotes::content')]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'additionalNotes::content')]")).clear();
	    browser.findElement(By.xpath("//*[contains(@id,'additionalNotes::content')]")).sendKeys(additionalNote);
	    Thread.sleep(4000);
	    }

	    if(!custSrcInspec.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'custSrcInspec::Label1')]")).click();
	    Thread.sleep(2000);
	    }
	    if(!govtSrcInspec.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'govtSrcInspec::Label1')]")).click();
	    Thread.sleep(2000);
	    }
	    if(!FAI.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'fai::Label1')]")).click();
	    Thread.sleep(3000);
	    }
	    if(!materialCert.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'materialCerts::Label1')]")).click();
	    }
	    if(!testReport.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'testReports::Label1')]")).click();
	    }
	    if(!dimInspec.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'dimensionalInspection::Label1')]")).click();
	    }
	    if(!FAA_Form.equals("NA"))
	    {
	    browser.findElement(By.xpath("//*[contains(@id,'a81303FaaForm::Label1')]")).click();
	    }
	    Thread.sleep(4000);
	    JavascriptExecutor js11 = (JavascriptExecutor)browser;
	      js11.executeScript("window.scrollBy(0, 2000)", "");
	    Thread.sleep(5000);
	   
	    browser.findElement(By.xpath("//button[contains(@id,'dEffAttr::ok')]")).click();
	    Thread.sleep(5000);
	    if(!shipTOContact.equals("NA"))
	      {
	      browser.findElement(By.xpath("//img[contains(@id, 'APRS1:sdi2::icon')]")).click();
	      Thread.sleep(3000);
	      WebElement action = browser.findElement(By.xpath("//*[contains(@id,'pgl10:LnsTbl::db')]/table/tbody/tr["+j+"]/td[2]/div/table/tbody/tr/td[31]"));
	      JavascriptExecutor jsre = (JavascriptExecutor)browser;
	      jsre.executeScript("arguments[0].scrollIntoView()", action);
	      Thread.sleep(3000);
	      action.click();
	      Thread.sleep(3000);
	      browser.findElement(By.xpath("//td[text()='Override Order Line']")).click();
	      Thread.sleep(3000);
	      browser.findElement(By.xpath("//input[contains(@id, 'APRS1:r5:0:shipToContactName1Id::content')]")).sendKeys(shipTOContact);
	      Thread.sleep(2000);
	      browser.findElement(By.xpath("//button[contains(@id, 'APRS1:r5:0:cb1')]")).click();
	   
	      }
	    CellStyle style = wb.createCellStyle();  
	    XSSFCell cell1 = sheet.getRow(i).createCell(67);
	    cell1.setCellValue("PASS");
	Font font = wb.createFont();
	font.setColor(IndexedColors.GREEN.getIndex());
	font.setBold(true);
	style = wb.createCellStyle();
	style.setFont(font);
	cell1.setCellStyle(style);
	fos = new FileOutputStream(srcFile);
	wb.write(fos);
	i++;
	    }
	    else
	    {
	    System.out.println("lines is not matched");
	    }
	 

	 
	}
	   
	   }
	}


	    }}

	    catch(Exception e)
	    {
	    e.printStackTrace();
	    }
	   }
	   @AfterTest
	   public void afterTest()
	   {
//	     browser.close();
	   }
	   
	   public static boolean isRowEmpty(Row row) {
	   for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	       Cell cell = row.getCell(c);
	       if (cell != null && cell.getCellType() != CellType.BLANK)
	           return false;
	   }
	   return true;
	 }
	   public void search() throws InterruptedException
	{
	    Thread.sleep(5000);
	 browser.findElement(By.linkText("Advanced")).click();
	 Thread.sleep(5000);
	 WebElement ele2 = browser.findElement(By.xpath("//select[contains(@id, 'operator1')]"));
	 Select element11 = new Select(ele2);
	 element11.selectByVisibleText("Equals");
	 Thread.sleep(2000);
	 browser.findElement(By.xpath("//input[contains(@id, 'value10')]")).sendKeys(orderNumber);
	 WebElement state = browser.findElement(By.xpath("//select[contains(@id, 'operator7::content')]"));
	 Select element1 = new Select(state);
	 element1.selectByVisibleText("Equals");
	 Thread.sleep(2000);
	 WebElement process = browser.findElement(By.xpath("//select[contains(@id, 'value70::content')]"));
	 Select element2 = new Select(process);
	 element2.selectByVisibleText("Processing");
	 Thread.sleep(3000);
	 browser.findElement(By.xpath("//button[contains(@id, 'search')]")).click();


	}
	   public void unitSellingPrice() throws InterruptedException
	{

	Thread.sleep(3000);
	 WebElement pencil= browser.findElement(By.xpath("(//img[contains(@id, 'i11:0:cil1::icon')])["+j+"]"));
	 JavascriptExecutor js = (JavascriptExecutor) browser;
	 js.executeScript("arguments[0].scrollIntoView();",pencil );
	pencil.click();
	 Thread.sleep(5000);
	 WebElement typ = browser.findElement(By.xpath("//select[contains(@id, 'APRS1:r15:0:t1:0:soc2::content')]"));
	 Select adjtype = new Select(typ);
	 adjtype.selectByVisibleText(type);
	 Thread.sleep(2000);
	 browser.findElement(By.xpath("//input[contains(@id, 'APRS1:r15:0:t1:0:it2::content')]")).sendKeys(pricingAdj);
	 WebElement reas = browser.findElement(By.xpath("//select[contains(@id, 'APRS1:r15:0:t1:0:soc1::content')]"));
	 Select resn = new Select(reas);
	 resn.selectByVisibleText(reason);
	 Thread.sleep(3000);
	 browser.findElement(By.xpath("//button[contains(@id, 'APRS1:mpa_dialog::ok')]")).click();
	 

	}
}
