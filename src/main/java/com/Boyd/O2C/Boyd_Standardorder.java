package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
//import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class Boyd_Standardorder {
	
	public static WebDriver browser;
	public static String Business_Unit;
	public static String Customer;
	public static String PurchaseOrder;
	public static String OrderType;
	public static String PrimarySalesperson;
	public static String Ship_to_Customer;
	public static String ship_to_Address;
	public static String Bill_to_Customer;
	public static String Bill_to_Address;
	public static String Payment_Terms;
	public static String Currency;
	public static String Currency_Convertion_Type;
	public static String Market_Segment;
	public static String CustomerAdvocate;
	public static String Ship_To_Contact;
	public static String Bill_To_Contact;
	public static String Customer_Email_Address;
	public static String PhoneNumber;
	public static String Print_ATO_Options;
	public static String QuoteNumber;
	public static String Warranty_Type;
	public static String Price_List;
	public static String Ship_Complete;
	public static String Loan_Requestor;
	public static String Hits_Order;
	public static String Project_Number;
	public static String Drop_Ship_Eligible;
	public static String Quote_Std_Margin;
	public static String Sales_Comp_StdMargin;
	public static String Opportunity_Type;
	public static String Demo_Drop_ship_Eligible;
	public static String Internal_Drop_Ship_Eligible;
	public static String IE_Drop_Ship_Eligible;
	public static String IFA_Drop_ship_Eligible;
	public static String Apply_Hold;
	public static String ShippingMethod_Header;
	public static String Requested_Date_Header;
	public static String Request_Type_Header;
	public static String FOB_Header;
	public static String FrieghtTerms_Header;
	public static String Shipping_Instructions_Header;
	public static String Warehouse_Header;
	public static String Demand_Class_Header;
	public static String Supplier_Header;

	public static String Item;
	public static String Quantity;
	public static String Price1;
	public static String Bundle_Part_Number;
	public static String FOB;
	public static String FrieghtTerms;
	public static String Shipping_Instructions;
	public static String ShippingMethod;
	public static String Requested_Date;
	public static String Request_Type;
	public static String Warehouse;
	public static String Demand_Class;
	public static String Supplier;

	public static int tabler;
	public static int count=0;
	public static int g=0;
	public static int h = 1;

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
	
//	browser.get("https://egmn-dev3.login.us2.oraclecloud.com/");
////	browser.get("https://egmn-dev4.fa.us2.oraclecloud.com/");
//	browser.findElement(By.id("userid")).sendKeys("Jiong.tang@harmonicinc.com");
//	browser.findElement(By.id("password")).sendKeys("welcome12345");
//	browser.findElement(By.id("btnActive")).click();
	
	
	browser.get("https://elme-dev2.fa.us8.oraclecloud.com/");
	browser.findElement(By.id("userid")).click();
	browser.findElement(By.id("userid")).sendKeys("Janakiram.Nalla");
	browser.findElement(By.id("password")).click();
	browser.findElement(By.id("password")).sendKeys("welcome1");
	browser.findElement(By.id("btnActive")).click();
	Thread.sleep(12000);
	
	browser.findElement(By.xpath("//a[text()='You have a new home page!']")).click();
	Thread.sleep(7000);
	browser.findElement(By.xpath("//*[text()='Order Management']")).click();
	Thread.sleep(3000);
	browser.findElement(By.id("itemNode_order_management_order_management_1")).click();
	Thread.sleep(3000);
	browser.findElement(By.xpath("//span[text()='Create Order']")).click();
	}
	@Test()
	public void Excel_data() throws Exception
	{
	File f = new File(System.getProperty("user.dir")+"\\Excel\\Standardorder.xlsx");
	FileInputStream fis = new FileInputStream(f);
	XSSFWorkbook wb = new XSSFWorkbook(fis);
	XSSFSheet sheet = wb.getSheetAt(0);
	sheet.getRow(0).createCell(57).setCellValue("Order");
	sheet.getRow(0).createCell(58).setCellValue("Result");
	sheet.getRow(0).createCell(59).setCellValue("Comments");
//	sheet.getRow(0).createCell(59).setCellValue("Final Status");
	int totalRows = sheet.getPhysicalNumberOfRows();
	System.out.println("Total number of Excel rows are :" +totalRows);
	
	if(sheet.getRow(1).getCell(58) == null)
	
	{
	for(int i =1;i<=totalRows;i++)
	{
	if(sheet.getRow(i) == null) {
	Thread.sleep(4000);
//	browser.findElement(By.xpath("//a[contains(@id, 'AP1:save::popEl')]")).click();
//	browser.findElement(By.xpath("//*[contains(@id, 'AP1:cmi2')]")).click();
	
	
	 browser.findElement(By.xpath("//span[text()='Submit']")).click();
	 WebDriverWait wait = new WebDriverWait(browser,350);
	 wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'AP1:cb14')]")));
	 browser.findElement(By.xpath("//*[contains(@id,'AP1:cb14')]")).click();
	 Thread.sleep(3000);
	 // browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:6:APVIEW1:SPb")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]")).click();
	break;
	}
	if(sheet.getRow(i) != null && isRowEmpty(sheet.getRow(i))) {
	Thread.sleep(4000);
//	browser.findElement(By.xpath("//a[contains(@id, 'AP1:save::popEl')]")).click();
//	browser.findElement(By.xpath("//*[contains(@id, 'AP1:cmi2')]")).click();
	
	
	 browser.findElement(By.xpath("//span[text()='Submit']")).click();
	 WebDriverWait wait = new WebDriverWait(browser,350);
	 wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'AP1:cb14')]")));
	 browser.findElement(By.xpath("//*[contains(@id,'AP1:cb14')]")).click();
	 Thread.sleep(3000);
	 // browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:6:APVIEW1:SPb")).click();
	 browser.findElement(By.xpath("//*[contains(@id,'APVIEW1:SPb')]")).click();
	Thread.sleep(5000);
	browser.findElement(By.xpath("//span[text()='Create Order']")).click();
	continue;
	}
	Business_Unit = sheet.getRow(i).getCell(0).getStringCellValue().trim();
	Customer = sheet.getRow(i).getCell(1).getStringCellValue().trim();
	PurchaseOrder = sheet.getRow(i).getCell(2).getStringCellValue().trim();
	OrderType = sheet.getRow(i).getCell(3).getStringCellValue().trim();
	PrimarySalesperson = sheet.getRow(i).getCell(4).getStringCellValue().trim();
	Ship_to_Customer = sheet.getRow(i).getCell(5).getStringCellValue();
	ship_to_Address = sheet.getRow(i).getCell(6).getStringCellValue().trim();
	Bill_to_Customer = sheet.getRow(i).getCell(7).getStringCellValue().trim();
	Bill_to_Address = sheet.getRow(i).getCell(8).getStringCellValue().trim();
	Payment_Terms = sheet.getRow(i).getCell(9).getStringCellValue().trim();
	Currency = sheet.getRow(i).getCell(10).getStringCellValue().trim();
	Currency_Convertion_Type = sheet.getRow(i).getCell(11).getStringCellValue().trim();
	Market_Segment = sheet.getRow(i).getCell(12).getStringCellValue().trim();
	CustomerAdvocate = sheet.getRow(i).getCell(13).getStringCellValue().trim();
	Ship_To_Contact = sheet.getRow(i).getCell(14).getStringCellValue().trim();
	Bill_To_Contact = sheet.getRow(i).getCell(15).getStringCellValue().trim();
	Customer_Email_Address = sheet.getRow(i).getCell(16).getStringCellValue().trim();
	PhoneNumber = sheet.getRow(i).getCell(17).getStringCellValue().trim();
	Print_ATO_Options = sheet.getRow(i).getCell(18).getStringCellValue().trim();
	QuoteNumber = sheet.getRow(i).getCell(19).getStringCellValue().trim();
	Warranty_Type = sheet.getRow(i).getCell(20).getStringCellValue().trim();
	Price_List = sheet.getRow(i).getCell(21).getStringCellValue().trim();
	Ship_Complete = sheet.getRow(i).getCell(22).getStringCellValue().trim();
	Loan_Requestor = sheet.getRow(i).getCell(23).getStringCellValue().trim();
	Hits_Order = sheet.getRow(i).getCell(24).getStringCellValue().trim();
	Project_Number = sheet.getRow(i).getCell(25).getStringCellValue().trim();
	Drop_Ship_Eligible = sheet.getRow(i).getCell(26).getStringCellValue().trim();
	Quote_Std_Margin = sheet.getRow(i).getCell(27).getStringCellValue().trim();
	Sales_Comp_StdMargin = sheet.getRow(i).getCell(28).getStringCellValue().trim();
	Opportunity_Type = sheet.getRow(i).getCell(29).getStringCellValue().trim();
	Demo_Drop_ship_Eligible = sheet.getRow(i).getCell(30).getStringCellValue().trim();
	Internal_Drop_Ship_Eligible = sheet.getRow(i).getCell(31).getStringCellValue().trim();
	IE_Drop_Ship_Eligible = sheet.getRow(i).getCell(32).getStringCellValue().trim();
	IFA_Drop_ship_Eligible = sheet.getRow(i).getCell(33).getStringCellValue().trim();
	

	Apply_Hold = sheet.getRow(i).getCell(34).getStringCellValue().trim();
	ShippingMethod_Header = sheet.getRow(i).getCell(35).getStringCellValue().trim();
	Requested_Date_Header = sheet.getRow(i).getCell(36).getStringCellValue().trim();
	Request_Type_Header = sheet.getRow(i).getCell(37).getStringCellValue().trim();
	FOB_Header = sheet.getRow(i).getCell(38).getStringCellValue().trim();
	FrieghtTerms_Header = sheet.getRow(i).getCell(39).getStringCellValue().trim();
	Shipping_Instructions_Header = sheet.getRow(i).getCell(40).getStringCellValue().trim();
	Warehouse_Header = sheet.getRow(i).getCell(41).getStringCellValue().trim();
	Demand_Class_Header = sheet.getRow(i).getCell(42).getStringCellValue().trim();
	Supplier_Header = sheet.getRow(i).getCell(43).getStringCellValue().trim();

	Item = sheet.getRow(i).getCell(44).getStringCellValue().trim();
	Quantity = sheet.getRow(i).getCell(45).getStringCellValue().trim();
	Price1 = sheet.getRow(i).getCell(46).getStringCellValue().trim();
	Bundle_Part_Number = sheet.getRow(i).getCell(47).getStringCellValue().trim();
	FOB = sheet.getRow(i).getCell(48).getStringCellValue().trim();
	FrieghtTerms = sheet.getRow(i).getCell(49).getStringCellValue().trim();
	Shipping_Instructions = sheet.getRow(i).getCell(50).getStringCellValue().trim();
	ShippingMethod = sheet.getRow(i).getCell(51).getStringCellValue().trim();
	Requested_Date = sheet.getRow(i).getCell(52).getStringCellValue().trim();
	Request_Type = sheet.getRow(i).getCell(53).getStringCellValue().trim();
	Warehouse = sheet.getRow(i).getCell(54).getStringCellValue().trim();
	Demand_Class = sheet.getRow(i).getCell(55).getStringCellValue().trim();
	Supplier = sheet.getRow(i).getCell(56).getStringCellValue().trim();
	// Apply_Hold = sheet.getRow(i).getCell(64).getStringCellValue().trim();



//	List<String> checkbox = new ArrayList<String>();
//	checkbox.add(OptionalItem1);
//	checkbox.add(OptionalItem2);
//	checkbox.add(OptionalItem3);
//	checkbox.add(optionalItem4);
//	checkbox.add(OptionalItem5);
//	checkbox.add(OptionalItem6);
//
//	List<String> input = new ArrayList<String>();
//	input.add(Quantity1);
//	input.add(Quantity2);
//	input.add(Quantity3);
//	input.add(Quantity4);
//	input.add(Quantity5);
//	input.add(Quantity6);

	if(!Customer.equals("NA") && !PurchaseOrder.equals("NA")) {
	     updateCustomerInfo();
	}
	   Thread.sleep(5000);
	   try {
	         WebElement search = browser.findElement(By.xpath("//*[contains(@id,'itemNumberId:searchIcoId::icon')]"));    
	         JavascriptExecutor js = (JavascriptExecutor)browser;
	         js.executeScript("arguments[0].scrollIntoView()", search);
	// search.click();
	if(search.isEnabled())
	{
	    highLightElement(browser,search);
	    Thread.sleep(3000);
	    JavascriptExecutor executor = (JavascriptExecutor)browser;
	    executor.executeScript("arguments[0].click();", search);
	    System.out.println("Search icon clicked");
	}
	else
	{
	   System.out.println("Search icon not enabled");
	}
	Thread.sleep(2000);
	browser.findElement(By.xpath("//*[text()='vanced']")).click();
	Thread.sleep(4000);
	Select scdp = new Select(browser.findElement(By.xpath("//*[contains(@id,'saveSearch::content')]")));
	scdp.selectByVisibleText("Application Default copy");
	Thread.sleep(4000);
	List<WebElement> Expand = browser.findElements(By.xpath("//*[contains(@title, 'Expand Advanced Search')]"));
	int expandCount = Expand.size();
	System.out.println("expandCount="+expandCount);
	if(expandCount>0)
	{
	WebElement element = browser.findElement(By.xpath("//*[contains(@id,'itemNumberId:Popup1:0:Advan1:0:efqrp::_afrDscl')]"));
	JavascriptExecutor executor = (JavascriptExecutor)browser;
	executor.executeScript("arguments[0].click();", element);
	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content')]")).sendKeys(Item);
	Thread.sleep(8000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp::search')]")).click();
	}
	else
	{
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content')]")).sendKeys(Item);
	Thread.sleep(8000);
	browser.findElement(By.xpath("//button[text()='Search']")).click();
	}
	Thread.sleep(8000);
	//**Problem in adding Lines**//
	try {
	     browser.findElement(By.xpath("//*[contains(@id, 'rstab:_ATp:table1::db')]/table/tbody/tr[1]/td[1]")).click();
	     Thread.sleep(5000);
	     browser.findElement(By.xpath("//*[contains(@id,'AP1:itemNumberId:cb1')]")).click();
	     Thread.sleep(8000);
	     WebElement Qunty = browser.findElement(By.xpath("//*[contains(@id, 'createLineQuantity::content')]"));
	     WebDriverWait wait = new WebDriverWait(browser,350);
	     wait.until(ExpectedConditions.elementToBeClickable(Qunty));
	     Qunty.click();
	     Qunty.clear();
	     Thread.sleep(3000);
	     Qunty.sendKeys(Quantity);
	}
	catch(Exception e)
	{
	   e.printStackTrace();
	   System.out.println("Unable to add Item and Quantity in Item");
	}
	//**Problem in adding Lines**//
	Thread.sleep(5000);
	}
	catch(Exception exe)
	{
	try {
		Thread.sleep(5000);
		browser.findElement(By.xpath("//button[contains(@id, 'dialogId::cancel')]")).click();
		Thread.sleep(3000);
	}
	catch(Exception e)
	{
       
	}
		System.out.println("Unable to add Item...."+Item);
		sheet.getRow(i).createCell(58).setCellValue("Fail");
		sheet.getRow(i).createCell(59).setCellValue("Item not available or Unable to add Item");
		Updatefile(f,wb);
		continue;
	}
	@SuppressWarnings("unused")
	boolean a = true;
	@SuppressWarnings("unused")
	String option4;
	
	//Code for Standard Items
	
		Thread.sleep(5000);
		WebElement addElement = browser.findElement(By.xpath("//span[text()='Add']"));
	if(addElement.isEnabled())
	{
		highLightElement(browser,addElement);
		Thread.sleep(3000);
		JavascriptExecutor executor = (JavascriptExecutor)browser;
		executor.executeScript("arguments[0].click();", addElement);
		System.out.println("Add button clicked");
	}
	else
	{
		System.out.println("Add button not enabled");
	}
	  Thread.sleep(6000);
	  JavascriptExecutor js = (JavascriptExecutor)browser;
	  js.executeScript("window.scrollBy(0, 2000)", "");
	  Thread.sleep(5000);
	  try {
	List<WebElement> tablerows2 = browser.findElements(By.xpath("//*[contains(@id,'AP1:pc1:t1::db')]/table[1]/tbody/tr"));
	Thread.sleep(5000);
	tabler = tablerows2.size();
	System.out.println("Size of table :" +tabler);
	  }
	  catch(Exception e)
	  {
		  List<WebElement> tablerows2 = browser.findElements(By.xpath("//*[contains(@id,'AP1:pc1:t1::db')]/table[2]/tbody/tr"));
		  Thread.sleep(5000);
		  tabler = tablerows2.size();
		  System.out.println("Size of table :" +tabler);
	  }
		   Thread.sleep(3000);
		   Actions act = new Actions(browser);
		   WebElement element = browser.findElement(By.xpath("//span[contains(text(), '"+Item+"')]"));
		   Thread.sleep(2000);
		   act.moveToElement(element).doubleClick().perform();
	 
	if(!Bundle_Part_Number.equals("NA") || !Price1.equals("NA"))
	{
		  WebElement ele = browser.findElement(By.xpath("//span[text()='"+Item+"']/../../../../../../../../../../../../../../..//button[contains(@title,'Actions')]"));
		  JavascriptExecutor jse = (JavascriptExecutor)browser;
		  jse.executeScript("arguments[0].scrollIntoView()", ele);
		  ele.click();
		  Thread.sleep(3000);
		  browser.findElement(By.xpath("//tr[contains(@id,'lineAdditionalInfo')]")).click();
		  Thread.sleep(4000);
		  browser.findElement(By.linkText("Global Data Elements")).click();
	if(!Bundle_Part_Number.equals("NA"))
	{
	     bundlenumber();
	}
			Thread.sleep(4000);
			browser.findElement(By.xpath("//a[text()='Pricing Additional Information']")).click();
			Thread.sleep(5000);
	if(!Price1.equals("NA"))
	{
	    price();
	}
	   browser.findElement(By.xpath("//*[contains(@id,'AP1:dEffAttr::ok')]")).click();
	}

		JavascriptExecutor jsvertical = (JavascriptExecutor)browser;
		jsvertical.executeScript("window.scrollBy(-1000,0)","");
		Thread.sleep(7000);

	//**Start Line level Update Line**//

	if(!FrieghtTerms.equals("NA") || !FOB.equals("NA") || !ShippingMethod.equals("NA") || !Requested_Date.equals("NA") || !Shipping_Instructions.equals("NA"))
	{
	    browser.findElement(By.xpath("//span[text()='Update Lines']")).click();
	if(!FrieghtTerms.equals("NA"))
	{
		browser.findElement(By.xpath("//*[text()='Freight Terms']")).click();
		browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
		Thread.sleep(4000);
	}
	if(!FOB.equals("NA"))
	{
		browser.findElement(By.xpath("//*[text()='FOB']")).click();
		browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
		Thread.sleep(4000);
	}
	// if(!Payment_Terms.equals("NA"))
	// {
	// browser.findElement(By.xpath("//*[text()='Payment Terms']")).click();
	// browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
	// Thread.sleep(4000);
	// }
	if(!ShippingMethod.equals("NA"))
	{
		browser.findElement(By.xpath("//*[text()='Shipping Method']")).click();
		browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
		Thread.sleep(4000);
	}
	if(!Requested_Date.equals("NA"))
	{
		browser.findElement(By.xpath("//*[text()='Requested Date']")).click();
		browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
		Thread.sleep(4000);
	}
	if(!Shipping_Instructions.equals("NA"))
	{
		browser.findElement(By.xpath("//*[text()='Shipping Instructions']")).click();
		browser.findElement(By.xpath("//a[contains(@title,'Move selected items to: Selected')]")).click();
		Thread.sleep(4000);
	}
	    browser.findElement(By.xpath("//*[text()='ext']")).click();
	if(!FrieghtTerms.equals("NA"))
	{
		Select dropdown = new Select(browser.findElement(By.xpath("(//select[contains(@class,'x2h')])[2]")));                                                      
		dropdown.selectByVisibleText(FrieghtTerms);
		Thread.sleep(2000);
	}
	if(!FOB.equals("NA"))
	{
	   Select fob1 = new Select(browser.findElement(By.xpath("//*[contains(@id,'SP2:t2:1:soc1::content')]")));
	   fob1.selectByVisibleText(FOB);
	   Thread.sleep(2000);
	}
	// if(!Payment_Terms.equals("NA"))
	// {
	// Select payterms = new Select(browser.findElement(By.xpath("//*[contains(@id,'SP2:t2:2:soc1::content')]")));
	// payterms.selectByVisibleText(Payment_Terms);
	// Thread.sleep(2000);
	// }
	if(!ShippingMethod.equals("NA"))
	{
		browser.findElement(By.xpath("//*[contains(@id,'integerValueId::content')]")).sendKeys(ShippingMethod);
		Thread.sleep(2000);
	}
	    Thread.sleep(3000);
	if(!Requested_Date.equals("NA"))
	{
		browser.findElement(By.xpath("//*[contains(@id,'SP2:t2:3:id2::content')]")).sendKeys(Requested_Date);
		Thread.sleep(2000);
//		Select type = new Select(browser.findElement(By.xpath("//*[contains(@id,'SP2:t2:5:soc1::content')]")));
		Select type = new Select(browser.findElement(By.xpath("(//*[contains(@id,'soc1::content')])[3]")));
		type.selectByVisibleText(Request_Type);
	}
	if(!Shipping_Instructions.equals("NA"))
	{
		browser.findElement(By.xpath("//*[contains(@id,'SP2:t2:4:it2::content')]")).sendKeys(Shipping_Instructions);
		Thread.sleep(2000);
	}

		Thread.sleep(5000);
		browser.findElement(By.xpath("//*[text()='ave and Close']")).click();
	}
	else
	{
		Thread.sleep(6000);
		browser.findElement(By.xpath("//span[text()='Update Lines']")).click();
		Thread.sleep(8000);
		WebDriverWait wait = new WebDriverWait(browser,350);
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(), 'ancel')]")));
		browser.findElement(By.xpath("//span[contains(text(), 'ancel')]")).click();
	}
	
		Thread.sleep(5000);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:sdi1::icon')]")).click();
		Thread.sleep(5000);
		System.out.println("Order got success");
		WebElement order = browser.findElement(By.xpath("//*[contains(@id,'AP1:SPph::_afrTtxt')]"));
		String so = order.getText();
		int ordernumber = getNumericValue(so);
		System.out.println("Order value :" +ordernumber);
		sheet.getRow(i).createCell(57).setCellValue(ordernumber);
		sheet.getRow(i).createCell(58).setCellValue("Pass");
		Updatefile(f,wb);
	}
	
	}
	else
	{
		System.out.println("File is already Processed");
	}
	
	
	
	
	
	try {
	   wb.close();
	  } catch(Exception e) {
	 
	  }
	}

	 private void updateCustomerInfo() throws Exception {
	Select bussinesunit = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:soc3::content')]")));
	bussinesunit.selectByVisibleText(Business_Unit);
	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:partyNameId::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:partyNameId::content')]")).sendKeys(Customer);
	Thread.sleep(4000);
	try {
	browser.findElement(By.xpath("//a[text()='More...']")).click();
	Thread.sleep(2000);
	List<WebElement> orderTab = browser.findElements(By.xpath("//*[contains(@id, 'AP1:resId1::db')]/table/tbody/tr"));
	int orderTable = orderTab.size();
	System.out.println("customers="+orderTable);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:resId1::db')]/table/tbody/tr["+orderTable+"]/td[1]")).click();
	Thread.sleep(3000);
	browser.findElement(By.xpath("//button[contains(@id, 'AP1:partyNameId::lovDialogId::ok')]")).click();
	}
	catch(Exception e)
	{
		System.out.println("Unable to find more button");
	}
	Thread.sleep(3000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:partyNameId::content')]")).sendKeys(Keys.ENTER);
	Thread.sleep(12000);
	
	if(!Ship_to_Customer.equals("NA"))
	{
		browser.findElement(By.xpath("//*[contains(@id,'AP1:shipToPartyNameId::lovIconId')]")).click();
		Thread.sleep(6000);
		browser.findElement(By.xpath("//*[contains(@id,'shipToPartyNameId::dropdownPopup::popupsearch')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:q2:value00::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'AP1:q2:value00::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'AP1:q2:value00::content')]")).sendKeys(Ship_to_Customer);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:q2::search')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'AP1:resId6::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'AP1:shipToPartyNameId::lovDialogId::ok')]")).click();
	}
	Thread.sleep(8000);
	browser.findElement(By.xpath("//a[contains(@title,'Search: Ship-to Address')]")).click();                      
	Thread.sleep(5000);
	browser.findElement(By.xpath("//a[text()='Search...']")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId2:value00::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId2:value00::content')]")).sendKeys(ship_to_Address);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:qryId2::search')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:resId4::db')]/table/tbody/tr/td[1]")).click();
	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:shipToAddress::lovDialogId::ok')]")).click();
	Thread.sleep(8000);
	browser.findElement(By.xpath("//*[contains(@id,'billToPartyNameId::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'billToPartyNameId::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'billToPartyNameId::content')]")).sendKeys(Bill_to_Customer);
	Thread.sleep(6000);
	browser.findElement(By.xpath("//a[contains(@id, 'billToPartyNameId::_afrautosuggestmorelink')]")).click();
	Thread.sleep(5000);

	try {
	browser.findElement(By.xpath("//*[contains(@id,'AP1:billToPartyNameId::lovDialogId::ok')]")).click();
	}
	catch(Exception e)
	{
	System.out.println("Ok Button not displayed.");
	}
	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:it1::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:it1::content')]")).sendKeys(PurchaseOrder);
	browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::content')]")).sendKeys(OrderType);
	try
	{
		browser.findElement(By.xpath("//*[text()='More...']")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::_afrLovInternalQueryId::search')]")).click();
	    Thread.sleep(3000);
	    browser.findElement(By.xpath("//*[contains(@id,'orderType1Id_afrLovInternalTableId::db')]/table/tbody/tr[1]/td[1]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::lovDialogId::ok')]")).click();
	 }
	catch(Exception e)
	{
		browser.findElement(By.xpath("//*[contains(@id,'orderType1Id::content')]")).sendKeys(Keys.ENTER);
	}
	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[contains(@id,'primarySalesPersonNameId::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'primarySalesPersonNameId::content')]")).sendKeys(PrimarySalesperson);
	Thread.sleep(8000);
	browser.findElement(By.xpath("//*[text()='Actions']")).click();
	browser.findElement(By.xpath("//td[text()='Edit Currency Details']")).click();
	Select curr = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc2::content')]")));
	curr.selectByVisibleText(Currency);
	browser.findElement(By.xpath("//*[contains(@id,'AccountingTypeLOV::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AccountingTypeLOV::content')]")).sendKeys(Currency_Convertion_Type);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:oc_dialog::ok')]")).click();
	Thread.sleep(8000);
//	browser.findElement(By.xpath("//*[text()='Actions']")).click();
//	browser.findElement(By.xpath("//*[text()='Edit Additional Information']")).click();
//	Thread.sleep(3000);
//	if(!Market_Segment.equals("NA"))
//	{
//	marketsegment();
//	}
//	if(!CustomerAdvocate.equals("NA"))
//	{
//	customeradvocate();
//	}
//	if(!Ship_To_Contact.equals("NA"))
//	{
//	shiptocontact();
//	}
//	if(!Bill_To_Contact.equals("NA"))
//	{
//	billtocontact();
//	}
//
//	if(!Customer_Email_Address.equals("NA"))
//	{
//	customeremailaddress();
//	}
//	if(!PhoneNumber.equals("NA"))
//	{
//	phonenumber();
//	}
//	if(!Print_ATO_Options.equals("NA"))
//	{
//	printatooptions();
//	}
//	if(!QuoteNumber.equals("NA"))
//	{
//	quotenumber();
//	}
//	if(!Warranty_Type.equals("NA"))
//	{
//	warranty();
//	}
//
//	if(!Price_List.equals("NA"))
//	{
//	WebElement dropdown = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOGlobalData:0:pricelist::lovIconId')]"));
//	JavascriptExecutor drop = (JavascriptExecutor)browser;
//	drop.executeScript("arguments[0].scrollIntoView()", dropdown);
//	dropdown.click();
//	Thread.sleep(5000);
//	browser.findElement(By.xpath("//a[text()='Search...']")).click();
//	pricelist();
//	browser.findElement(By.xpath("//*[contains(@id,'pricelist::_afrLovInternalQueryId::search')]")).click();
//	browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOGlobalData:0:pricelist_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
//	browser.findElement(By.xpath("//*[contains(@id,'pricelist::lovDialogId::ok')]")).click();
//	}
//	Thread.sleep(5000);
//	if(!Ship_Complete.equals("NA"))
//	{
//	shipcomplete();
//	}
//	Thread.sleep(5000);
//	if(!Loan_Requestor.equals("NA"))
//	{
//		loan_requestor();
//	}
//	Thread.sleep(3000);
//	if(!Hits_Order.equals("NA"))
//	{
//	HitsOrder();
//	}
//	Thread.sleep(5000);
//	browser.findElement(By.linkText("Project Shipment")).click();
//	if(!Quote_Std_Margin.equals("NA"))
//	{
//	quotestdmargin();
//	}
//	Thread.sleep(5000);
//	browser.findElement(By.linkText("Project Shipment")).click();
//	if(!Project_Number.equals("NA"))
//	{
//	projectnumber();
//	}
//	Thread.sleep(3000);
//	if(!Drop_Ship_Eligible.equals("NA"))
//	{
//		dropship_eligible();
//	}
//	if(!Sales_Comp_StdMargin.equals("NA"))
//	{
//	salescompstdmargin();
//	}
//	if(!Opportunity_Type.equals("NA"))
//	{
//	opportunitytype();
//	}
//	Thread.sleep(5000);
//	browser.findElement(By.linkText("Demo Qualification")).click();
//	if(!Demo_Drop_ship_Eligible.equals("NA"))
//	{
//		dropship_Demoqualification();
//	}
//	Thread.sleep(5000);
//	browser.findElement(By.linkText("Internal")).click();
//	if(!Internal_Drop_Ship_Eligible.equals("NA"))
//	{
//		dropship_internal();
//	}
//	Thread.sleep(5000);
//	browser.findElement(By.linkText("Internal Expensed")).click();
//	if(!IE_Drop_Ship_Eligible.equals("NA"))
//	{
//		dropship_internalexpensed();
//	}
//	Thread.sleep(5000);
//	browser.findElement(By.linkText("Internal Fixed Asset")).click();
//	if(!IFA_Drop_ship_Eligible.equals("NA"))
//	{
//		dropship_internalfixedasset();
//	}
//	
//	Thread.sleep(5000);
//	browser.findElement(By.xpath("//*[contains(@id,'AP1:dEffAttr::ok')]")).click();
//
//	//**Apply Hold code added**//
//	Thread.sleep(4000);
//	if(!Apply_Hold.equals("NA"))
//	{
//	browser.findElement(By.xpath("(//a[contains(text(), 'Actions')])[1]")).click();
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//*[contains(@id,'AP1:m4')]/td[2]")).click();
//	Thread.sleep(2000);
//	browser.findElement(By.xpath("//*[contains(@id,'AP1:cmi6')]/td[2]")).click();
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//input[contains(@id, 'holdNameId::content')]")).click();
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//a[contains(text(), 'Search...')]")).click();
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//input[contains(@id, 'value00::content')]")).sendKeys(Apply_Hold);
//	browser.findElement(By.xpath("//button[contains(@id, 'search')]")).click();
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//*[contains(@id,'holdNameId_afrLovInternalTableId')]/table/tbody/tr[1]/td[1]")).click();
//	Thread.sleep(2000);
//	browser.findElement(By.xpath("//button[contains(@id, 'lovDialogId::ok')]")).click();
//	Thread.sleep(2000);
//	browser.findElement(By.xpath("//button[contains(text(), 'ave and Close')]")).click();
//
//	}


//	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:sdi3::icon')]")).click();
	Thread.sleep(8000);
	Select ba = new Select(browser.findElement(By.xpath("//*[contains(@id,'billToLocation::content')]")));
	ba.selectByVisibleText(Bill_to_Address);
//	browser.findElement(By.xpath("//*[contains(@id,'billToLocation::content')]")).click();
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//*[contains(@id,'billToLocation::dropdownPopup::popupsearch')]")).click();
//	browser.findElement(By.xpath("//*[contains(@id,'qryId2:value00::content')]")).click();
//	browser.findElement(By.xpath("//*[contains(@id,'qryId2:value00::content')]")).clear();
//	browser.findElement(By.xpath("//*[contains(@id,'qryId2:value00::content')]")).sendKeys(Bill_to_Address);
//	Thread.sleep(3000);
//	browser.findElement(By.xpath("//*[contains(@id,'qryId2::search')]")).click();
//	Thread.sleep(25000);
//	WebDriverWait wait = new WebDriverWait(browser,300);
//	wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'resId2::db')]/table/tbody/tr/td[1]")));
//	browser.findElement(By.xpath("//*[contains(@id,'resId2::db')]/table/tbody/tr/td[1]")).click();
//	Thread.sleep(4000);
//	browser.findElement(By.xpath("//*[contains(@id,'billToLocation::lovDialogId::ok')]")).click();
	Thread.sleep(6000);
	Select pt = new Select(browser.findElement(By.xpath("//*[contains(@id,'paymentTermId::content')]")));
	pt.selectByVisibleText(Payment_Terms);
	Thread.sleep(10000);

	//** Header Shipment details updated **//
//	WebDriverWait wait1 = new WebDriverWait(browser,300);
//	wait1.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[contains(@id,'AP1:sdi2::icon')]")));
	browser.findElement(By.xpath("//*[contains(@id,'AP1:sdi2::icon')]")).click();
	if(!ShippingMethod_Header.equals("NA"))
	{
	browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:shipMethodId::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:shipMethodId::content')]")).clear();
	Thread.sleep(4000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:shipMethodId::content')]")).sendKeys(ShippingMethod_Header);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:shipMethodId::content')]")).sendKeys(Keys.ENTER);
	}
	Thread.sleep(6000);
	if(!Requested_Date_Header.equals("NA"))
	{
	browser.findElement(By.xpath("//input[contains(@id, 'id1::content')]")).click();
	browser.findElement(By.xpath("//input[contains(@id, 'id1::content')]")).clear();
	Thread.sleep(4000);
	browser.findElement(By.xpath("//input[contains(@id, 'id1::content')]")).sendKeys(Requested_Date_Header);
	}
	Thread.sleep(5000);
	if(!Request_Type_Header.equals("NA"))
	{
	WebElement shipOn = browser.findElement(By.xpath("//select[contains(@id, 'soc1::content')]"));
	Select reqType = new Select(shipOn);
	reqType.selectByVisibleText(Request_Type_Header);
	}
	Thread.sleep(8000);
	browser.findElement(By.linkText("Shipping")).click();
	if(!FOB_Header.equals("NA"))
	{
		Thread.sleep(5000);
	Select fob = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc3::content')]")));
	fob.selectByVisibleText(FOB_Header);
	Thread.sleep(5000);
	}
	if(!FrieghtTerms_Header.equals("NA"))
	{
	Select FT = new Select(browser.findElement(By.xpath("//*[contains(@id,'AP1:r5:0:soc4::content')]")));
	FT.selectByVisibleText(FrieghtTerms_Header);
	}
	if(!Shipping_Instructions_Header.equals("NA"))
	{
		Thread.sleep(5000);
	browser.findElement(By.xpath("//textarea[contains(@id, 'it2::content')]")).click();
	browser.findElement(By.xpath("//textarea[contains(@id, 'it2::content')]")).clear();
	Thread.sleep(4000);
	browser.findElement(By.xpath("//textarea[contains(@id, 'it2::content')]")).sendKeys(Shipping_Instructions_Header);
	}
	Thread.sleep(6000);
	browser.findElement(By.linkText("Supply")).click();
	Thread.sleep(5000);
	if(!Warehouse_Header.equals("NA"))
	{
	warehousedetails();
	}
	Thread.sleep(4000);
	if(!Demand_Class_Header.equals("NA"))
	{
	demandclass();
	}
	Thread.sleep(4000);
	browser.findElement(By.xpath("//*[contains(@id,'AP1:sdi1::icon')]")).click();        
	}

	public void billtocontact()
	{
	browser.findElement(By.xpath("//*[contains(@id,'billtocontact::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'billtocontact::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'billtocontact::content')]")).sendKeys(Bill_To_Contact);
	}

	public void marketsegment()
	{
	browser.findElement(By.xpath("//*[contains(@id,'marketsegment::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'marketsegment::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'marketsegment::content')]")).sendKeys(Market_Segment);
	}
	public void customeremailaddress()
	{
	browser.findElement(By.xpath("//*[contains(@id,'customeremailaddress::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'customeremailaddress::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'customeremailaddress::content')]")).sendKeys(Customer_Email_Address);
	}
	public void phonenumber()
	{
	browser.findElement(By.xpath("//*[contains(@id,'phonenumber::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'phonenumber::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'phonenumber::content')]")).sendKeys(PhoneNumber);
	}
	public void quotenumber()
	{
	browser.findElement(By.xpath("//*[contains(@id,'quotenumber::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'quotenumber::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'quotenumber::content')]")).sendKeys(QuoteNumber);
	}
	public void quotestdmargin()
	{
	browser.findElement(By.xpath("//*[contains(@id,'quotestdmargin::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'quotestdmargin::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'quotestdmargin::content')]")).sendKeys(Quote_Std_Margin);
	}
	public void salescompstdmargin()
	{
	browser.findElement(By.xpath("//*[contains(@id,'salescompstdmargin::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'salescompstdmargin::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'salescompstdmargin::content')]")).sendKeys(Sales_Comp_StdMargin);
	}
	public void HitsOrder()
	{
	browser.findElement(By.xpath("//*[contains(@id,'hitsOrder::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'hitsOrder::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'hitsOrder::content')]")).sendKeys(Hits_Order);
	}
	public void customeradvocate()
	{
	browser.findElement(By.xpath("//*[contains(@id,'customeradvocate::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'customeradvocate::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'customeradvocate::content')]")).sendKeys(CustomerAdvocate);
	}
	public void shiptocontact()
	{
	browser.findElement(By.xpath("//*[contains(@id,'shiptocontact::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'shiptocontact::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'shiptocontact::content')]")).sendKeys(Ship_To_Contact);
	}
	public void printatooptions()
	{
	browser.findElement(By.xpath("//*[contains(@id,'printatooptions::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'printatooptions::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'printatooptions::content')]")).sendKeys(Print_ATO_Options);
	}
	public void warranty()
	{
	browser.findElement(By.xpath("//*[contains(@id,'warrantytype::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'warrantytype::content')]")).clear();
	         browser.findElement(By.xpath("//*[contains(@id,'warrantytype::content')]")).sendKeys(Warranty_Type);
	}
	public void pricelist()
	{
	browser.findElement(By.xpath("//*[contains(@id,'pricelist::_afrLovInternalQueryId:value00::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'pricelist::_afrLovInternalQueryId:value00::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'pricelist::_afrLovInternalQueryId:value00::content')]")).sendKeys(Price_List);
	}
	public void shipcomplete()
	{
	browser.findElement(By.xpath("//*[contains(@id,'ShipComplete::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'ShipComplete::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'ShipComplete::content')]")).sendKeys(Ship_Complete);
	}
	
	public void loan_requestor()
	{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOGlobalData:0:loanerRequestor::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOGlobalData:0:loanerRequestor::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOGlobalData:0:loanerRequestor::content')]")).sendKeys(Loan_Requestor);
	}
	
	public void projectnumber()
	{
	browser.findElement(By.xpath("//*[contains(@id,'projectnumber::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'projectnumber::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'projectnumber::content')]")).sendKeys(Project_Number);
	}
	
	public void dropship_eligible() throws Exception
	{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::lovIconId')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::dropdownPopup::popupsearch')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).sendKeys(Drop_Ship_Eligible);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::_afrLovInternalQueryId::search')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOProjectShipment:0:dropshipeligible::lovDialogId::ok')]")).click();
	}
	
	public void dropship_Demoqualification() throws Exception
	{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::lovIconId')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::dropdownPopup::popupsearch')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).sendKeys(Demo_Drop_ship_Eligible);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::_afrLovInternalQueryId::search')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVODemoQualification:0:dropshipeligible::lovDialogId::ok')]")).click();
	}
	public void dropship_internal() throws Exception
	{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::lovIconId')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::dropdownPopup::popupsearch')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).sendKeys(Internal_Drop_Ship_Eligible);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::_afrLovInternalQueryId::search')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternal:0:dropshipeligible::lovDialogId::ok')]")).click();
	}
	public void dropship_internalexpensed() throws Exception
	{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::lovIconId')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::dropdownPopup::popupsearch')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).sendKeys(IE_Drop_Ship_Eligible);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::_afrLovInternalQueryId::search')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalExpensed:0:dropshipeligible::lovDialogId::ok')]")).click();
	}
	public void dropship_internalfixedasset() throws Exception
	{
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::lovIconId')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::dropdownPopup::popupsearch')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).clear();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::_afrLovInternalQueryId:value00::content')]")).sendKeys(IFA_Drop_ship_Eligible);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::_afrLovInternalQueryId::search')]")).click();
		Thread.sleep(3000);
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible_afrLovInternalTableId::db')]/table/tbody/tr/td[1]")).click();
		browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_HeaderEffDooHeadersAddInfoprivateVOInternalFixedAsset:0:dropshipeligible::lovDialogId::ok')]")).click();
	}
	
	
	public void opportunitytype()
	{
	browser.findElement(By.xpath("//*[contains(@id,'opportunitytype::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'opportunitytype::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'opportunitytype::content')]")).sendKeys(Opportunity_Type);
	}
	public void bundlenumber()
	{
	browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).sendKeys(Bundle_Part_Number);
	}
	public void price() throws Exception
	{
	WebElement Price = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOPricingAdditionalInformation:0:unitsellingprice::content')]"));
	Price.click();
	Price.clear();
	Thread.sleep(4000);
	Price.sendKeys(Price1);
	}
//	public void pricevalue() throws Exception
//	{
//	WebElement prc = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOPricingAdditionalInformation:0:unitsellingprice::content')]"));
//	prc.click();
//	        prc.clear();
//	Thread.sleep(4000);
//	prc.sendKeys(Price2);
//	}

//	public void pricevalue1() throws Exception
//	{
//	WebElement price = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOPricingAdditionalInformation:0:unitsellingprice::content')]"));
//	price.click();
//	        price.clear();
//	Thread.sleep(4000);
//	price.sendKeys(Price3);
//	}
//	public void pricevalue2() throws Exception
//	{
//	WebElement price = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOPricingAdditionalInformation:0:unitsellingprice::content')]"));
//	price.click();
//	        price.clear();
//	Thread.sleep(4000);
//	price.sendKeys(price4);
//	}
//	public void pricevalue3() throws Exception
//	{
//	WebElement price = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOPricingAdditionalInformation:0:unitsellingprice::content')]"));
//	price.click();
//	        price.clear();
//	Thread.sleep(4000);
//	price.sendKeys(price5);
//	}
//	public void pricevalue4() throws Exception
//	{
//	WebElement price = browser.findElement(By.xpath("//*[contains(@id,'CTXRNj_FulfillLineEffDooFulfillLinesAddInfoprivateVOPricingAdditionalInformation:0:unitsellingprice::content')]"));
//	price.click();
//	        price.clear();
//	Thread.sleep(4000);
//	price.sendKeys(price6);
//	}
	public void warehousedetails()
	{
	browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]")).click();
	browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]")).clear();
	browser.findElement(By.xpath("//*[contains(@id,'warehouseNameId::content')]")).sendKeys(Warehouse_Header);
	}
	public void demandclass()
	{
	Select demandclass = new Select(browser.findElement(By.xpath("//*[contains(@id,'soc6::content')]")));
	demandclass.selectByVisibleText(Demand_Class_Header);
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
	public void Quit_Browser()
	{
	// browser.quit();
	}
	public static void highLightElement(WebDriver browser,WebElement ele)
	{
	try {  
	           JavascriptExecutor js = (JavascriptExecutor) browser;  
	           js.executeScript("arguments[0].style.border='4px groove red'", ele);
	           Thread.sleep(1000);  
	           js.executeScript("arguments[0].style.border=''", ele);  
	      } catch (Exception e) {  
	           System.out.println(e);  
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
