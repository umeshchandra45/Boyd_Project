package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
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
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class Boyd_Updatebundle_Part {
	
	public static WebDriver browser;
	public static String Order_Number;
	public static String LineNumber;
	public static String status;
	public static String SalesOrderItem;
	public static String ATOModel_Qnty;
	public static String AddingOptionalItem1;
	public static String Quantity1;
	public static String ModelPrice1;
	public static String UnitSellingPrice1;
	public static String AddingOptionalItem2;
	public static String Quantity2;
	public static String ModelPrice2;
	public static String UnitSellingPrice2;
	public static String AddingOptionalItem3;
	public static String Quantity3;
	public static String ModelPrice3;
	public static String UnitSellingPrice3;
	public static String AddingOptionalItem4;
	public static String Quantity4;
	public static String ModelPrice4;
	public static String UnitSellingPrice4;
	public static String AddingOptionalItem5;
	public static String Quantity5;
	public static String ModelPrice5;
	public static String UnitSellingPrice5;
	public static String AddingOptionalItem6;
	public static String Quantity6;
	public static String ModelPrice6;
	public static String UnitSellingPrice6;
	public static String AddingOptionalItem7;
	public static String Quantity7;
	public static String ModelPrice7;
	public static String UnitSellingPrice7;
	public static String AddingOptionalItem8;
	public static String Quantity8;
	public static String ModelPrice8;
	public static String UnitSellingPrice8;
	public static String AddingOptionalItem9;
	public static String Quantity9;
	public static String ModelPrice9;
	public static String UnitSellingPrice9;
	public static String AddingOptionalItem10;
	public static String Quantity10;
	public static String ModelPrice10;
	public static String UnitSellingPrice10;
	public static String AddingOptionalItem11;
	public static String Quantity11;
	public static String ModelPrice11;
	public static String UnitSellingPrice11;
	public static String AddingOptionalItem12;
	public static String Quantity12;
	public static String ModelPrice12;
	public static String UnitSellingPrice12;
	public static String AddingOptionalItem13;
	public static String Quantity13;
	public static String ModelPrice13;
	public static String UnitSellingPrice13;
	public static String AddingOptionalItem14;
	public static String Quantity14;
	public static String ModelPrice14;
	public static String UnitSellingPrice14;
	public static String AddingOptionalItem15;
	public static String Quantity15;
	public static String ModelPrice15;
	public static String UnitSellingPrice15;
	public static String AddingOptionalItem16;
	public static String Quantity16;
	public static String ModelPrice16;
	public static String UnitSellingPrice16;
	public static String AddingOptionalItem17;
	public static String Quantity17;
	public static String ModelPrice17;
	public static String UnitSellingPrice17;
	public static String AddingOptionalItem18;
	public static String Quantity18;
	public static String ModelPrice18;
	public static String UnitSellingPrice18;
	public static String AddingOptionalItem19;
	public static String Quantity19;
	public static String ModelPrice19;
	public static String UnitSellingPrice19;
	public static String AddingOptionalItem20;
	public static String Quantity20;
	public static String ModelPrice20;
	public static String UnitSellingPrice20;
	public static String AddingOptionalItem21;
	public static String Quantity21;
	public static String ModelPrice21;
	public static String UnitSellingPrice21;
	public static String AddingOptionalItem22;
	public static String Quantity22;
	public static String ModelPrice22;
	public static String UnitSellingPrice22;
	public static String AddingOptionalItem23;
	public static String Quantity23;
	public static String ModelPrice23;
	public static String UnitSellingPrice23;
	public static String AddingOptionalItem24;
	public static String Quantity24;
	public static String ModelPrice24;
	public static String UnitSellingPrice24;
	public static String AddingOptionalItem25;
	public static String Quantity25;
	public static String ModelPrice25;
	public static String UnitSellingPrice25;
	public static String AddingOptionalItem26;
	public static String Quantity26;
	public static String ModelPrice26;
	public static String UnitSellingPrice26;
	public static String AddingOptionalItem27;
	public static String Quantity27;
	public static String ModelPrice27;
	public static String UnitSellingPrice27;
	public static String AddingOptionalItem28;
	public static String Quantity28;
	public static String ModelPrice28;
	public static String UnitSellingPrice28;
	public static String AddingOptionalItem29;
	public static String Quantity29;
	public static String ModelPrice29;
	public static String UnitSellingPrice29;
	public static String AddingOptionalItem30;
	public static String Quantity30;
	public static String ModelPrice30;
	public static String UnitSellingPrice30;
	public static String optionClass1;
	public static String optionClass2;
	public static String optionClass3;
	public static String Bundle_Part_Number;


	   public static String cancelItem;
	public static String cancelReason;
	public static int k=1;
	public static String Status;
	public static int tablesize;
	public static int n=1;
	public static int tabler;
	FileInputStream fis;
	FileOutputStream fos;
	File f;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	int i;

	// public static String linevalue;
	// public static int i;

	@BeforeTest()
	public void Login_Page() throws Exception
	{
	WebDriverManager.chromedriver().setup();
	ChromeOptions options = new ChromeOptions();
	options.setPageLoadStrategy(PageLoadStrategy.NONE);
	browser = new ChromeDriver(options);
	browser.manage().window().maximize();
	browser.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
	browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
	browser.get("https://egmn-dev4.login.us2.oraclecloud.com/");
	WebElement username = browser.findElement(By.id("userid"));
	highLightElement(browser,username);
	username.sendKeys("Jiong.tang@harmonicinc.com");
	WebElement password = browser.findElement(By.id("password"));
	highLightElement(browser,password);
	password.sendKeys("welcome12345");
	WebElement action = browser.findElement(By.id("btnActive"));
	highLightElement(browser,action);
	action.click();
	Thread.sleep(45000);
	WebElement button = browser.findElement(By.xpath("//a[text()='You have a new home page!']"));
	highLightElement(browser,button);
	button.click();
	Thread.sleep(15000);
	WebElement button1 = browser.findElement(By.xpath("//*[text()='Order Management']"));
	highLightElement(browser,button1);
	button1.click();
	WebElement button2 = browser.findElement(By.id("itemNode_order_management_order_management_1"));
	highLightElement(browser, button2);
	button2.click();
	Thread.sleep(9000);
	// WebElement search = browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:0:AP1:qq1:it1::content\"]"));
	// highLightElement(browser, search);
	// search.click();
	}
	@Test
	public void Home_Page() throws Exception
	{
	f = new File(System.getProperty("user.dir")+"\\Excel\\HRM-Bundle part Automation Template.xlsx");
	fis = new FileInputStream(f);
	wb = new XSSFWorkbook(fis);
	sheet = wb.getSheetAt(0);
	int totalrows = sheet.getPhysicalNumberOfRows();
	System.out.println("Total number of rows from Excel is :" +totalrows);
	for(i=1; i<=totalrows;)
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
	Order_Number = sheet.getRow(i).getCell(0).getStringCellValue().trim();
	LineNumber = sheet.getRow(i).getCell(1).getStringCellValue().trim();
	status = sheet.getRow(i).getCell(2).getStringCellValue().trim();
	SalesOrderItem = sheet.getRow(i).getCell(3).getStringCellValue().trim();
	cancelItem = sheet.getRow(i).getCell(128).getStringCellValue().trim();
	cancelReason = sheet.getRow(i).getCell(129).getStringCellValue().trim();
	System.out.println("Excel line value is :" +LineNumber);

	if(!Order_Number.equals("NA"))
	{
	Search();
	}

	if(!cancelItem.equals("NA"))
	    {
	    cancellItem();
	    i++;
	    continue;
	   
	    }

	if(LineNumber.equalsIgnoreCase("NA"))
	{
	SalesOrderItem = sheet.getRow(i).getCell(3).getStringCellValue().trim();
	ATOModel_Qnty = sheet.getRow(i).getCell(4).getStringCellValue().trim();

	optionClass1 = sheet.getRow(i).getCell(5).getStringCellValue().trim();

	AddingOptionalItem1 = sheet.getRow(i).getCell(6).getStringCellValue().trim();
	Quantity1 = sheet.getRow(i).getCell(7).getStringCellValue().trim();
	ModelPrice1 = sheet.getRow(i).getCell(8).getStringCellValue().trim();
	UnitSellingPrice1 = sheet.getRow(i).getCell(9).getStringCellValue().trim();
	AddingOptionalItem2 = sheet.getRow(i).getCell(10).getStringCellValue().trim();
	Quantity2 = sheet.getRow(i).getCell(11).getStringCellValue().trim();
	ModelPrice2 = sheet.getRow(i).getCell(12).getStringCellValue().trim();
	UnitSellingPrice2 = sheet.getRow(i).getCell(13).getStringCellValue().trim();
	AddingOptionalItem3 = sheet.getRow(i).getCell(14).getStringCellValue().trim();
	Quantity3 = sheet.getRow(i).getCell(15).getStringCellValue().trim();
	ModelPrice3 = sheet.getRow(i).getCell(16).getStringCellValue().trim();
	UnitSellingPrice3 = sheet.getRow(i).getCell(17).getStringCellValue().trim();
	AddingOptionalItem4 = sheet.getRow(i).getCell(18).getStringCellValue().trim();
	Quantity4 = sheet.getRow(i).getCell(19).getStringCellValue().trim();
	ModelPrice4 = sheet.getRow(i).getCell(20).getStringCellValue().trim();
	UnitSellingPrice4 = sheet.getRow(i).getCell(21).getStringCellValue().trim();
	AddingOptionalItem5 = sheet.getRow(i).getCell(22).getStringCellValue().trim();
	Quantity5 = sheet.getRow(i).getCell(23).getStringCellValue().trim();
	ModelPrice5 = sheet.getRow(i).getCell(24).getStringCellValue().trim();
	UnitSellingPrice5 = sheet.getRow(i).getCell(25).getStringCellValue().trim();
	AddingOptionalItem6 = sheet.getRow(i).getCell(26).getStringCellValue().trim();
	Quantity6 = sheet.getRow(i).getCell(27).getStringCellValue().trim();
	ModelPrice6 = sheet.getRow(i).getCell(28).getStringCellValue().trim();
	UnitSellingPrice6 = sheet.getRow(i).getCell(29).getStringCellValue().trim();
	AddingOptionalItem7 = sheet.getRow(i).getCell(30).getStringCellValue().trim();
	Quantity7 = sheet.getRow(i).getCell(31).getStringCellValue().trim();
	ModelPrice7 = sheet.getRow(i).getCell(32).getStringCellValue().trim();
	UnitSellingPrice7 = sheet.getRow(i).getCell(33).getStringCellValue().trim();
	AddingOptionalItem8 = sheet.getRow(i).getCell(34).getStringCellValue().trim();
	Quantity8 = sheet.getRow(i).getCell(35).getStringCellValue().trim();
	ModelPrice8 = sheet.getRow(i).getCell(36).getStringCellValue().trim();
	UnitSellingPrice8 = sheet.getRow(i).getCell(37).getStringCellValue().trim();
	AddingOptionalItem9 = sheet.getRow(i).getCell(38).getStringCellValue().trim();
	Quantity9 = sheet.getRow(i).getCell(39).getStringCellValue().trim();
	ModelPrice9 = sheet.getRow(i).getCell(40).getStringCellValue().trim();
	UnitSellingPrice9 = sheet.getRow(i).getCell(41).getStringCellValue().trim();
	AddingOptionalItem10 = sheet.getRow(i).getCell(42).getStringCellValue().trim();
	Quantity10 = sheet.getRow(i).getCell(43).getStringCellValue().trim();
	ModelPrice10 = sheet.getRow(i).getCell(44).getStringCellValue().trim();
	UnitSellingPrice10 = sheet.getRow(i).getCell(45).getStringCellValue().trim();

	optionClass2 = sheet.getRow(i).getCell(46).getStringCellValue().trim();

	AddingOptionalItem11 = sheet.getRow(i).getCell(47).getStringCellValue().trim();
	Quantity11 = sheet.getRow(i).getCell(48).getStringCellValue().trim();
	ModelPrice11 = sheet.getRow(i).getCell(49).getStringCellValue().trim();
	UnitSellingPrice11 = sheet.getRow(i).getCell(50).getStringCellValue().trim();
	AddingOptionalItem12 = sheet.getRow(i).getCell(51).getStringCellValue().trim();
	Quantity12 = sheet.getRow(i).getCell(52).getStringCellValue().trim();
	ModelPrice12 = sheet.getRow(i).getCell(53).getStringCellValue().trim();
	UnitSellingPrice12 = sheet.getRow(i).getCell(54).getStringCellValue().trim();
	AddingOptionalItem13 = sheet.getRow(i).getCell(55).getStringCellValue().trim();
	Quantity13 = sheet.getRow(i).getCell(56).getStringCellValue().trim();
	ModelPrice13 = sheet.getRow(i).getCell(57).getStringCellValue().trim();
	UnitSellingPrice13 = sheet.getRow(i).getCell(58).getStringCellValue().trim();
	AddingOptionalItem14 = sheet.getRow(i).getCell(59).getStringCellValue().trim();
	Quantity14 = sheet.getRow(i).getCell(60).getStringCellValue().trim();
	ModelPrice14 = sheet.getRow(i).getCell(61).getStringCellValue().trim();
	UnitSellingPrice14 = sheet.getRow(i).getCell(62).getStringCellValue().trim();
	AddingOptionalItem15 = sheet.getRow(i).getCell(63).getStringCellValue().trim();
	Quantity15 = sheet.getRow(i).getCell(64).getStringCellValue().trim();
	ModelPrice15 = sheet.getRow(i).getCell(65).getStringCellValue().trim();
	UnitSellingPrice15 = sheet.getRow(i).getCell(66).getStringCellValue().trim();
	AddingOptionalItem16 = sheet.getRow(i).getCell(67).getStringCellValue().trim();
	Quantity16 = sheet.getRow(i).getCell(68).getStringCellValue().trim();
	ModelPrice16 = sheet.getRow(i).getCell(69).getStringCellValue().trim();
	UnitSellingPrice16 = sheet.getRow(i).getCell(70).getStringCellValue().trim();
	AddingOptionalItem17 = sheet.getRow(i).getCell(71).getStringCellValue().trim();
	Quantity17 = sheet.getRow(i).getCell(72).getStringCellValue().trim();
	ModelPrice17 = sheet.getRow(i).getCell(73).getStringCellValue().trim();
	UnitSellingPrice17 = sheet.getRow(i).getCell(74).getStringCellValue().trim();
	AddingOptionalItem18 = sheet.getRow(i).getCell(75).getStringCellValue().trim();
	Quantity18 = sheet.getRow(i).getCell(76).getStringCellValue().trim();
	ModelPrice18 = sheet.getRow(i).getCell(77).getStringCellValue().trim();
	UnitSellingPrice18 = sheet.getRow(i).getCell(78).getStringCellValue().trim();
	AddingOptionalItem19 = sheet.getRow(i).getCell(79).getStringCellValue().trim();
	Quantity19 = sheet.getRow(i).getCell(80).getStringCellValue().trim();
	ModelPrice19 = sheet.getRow(i).getCell(81).getStringCellValue().trim();
	UnitSellingPrice19 = sheet.getRow(i).getCell(82).getStringCellValue().trim();
	AddingOptionalItem20 = sheet.getRow(i).getCell(83).getStringCellValue().trim();
	Quantity20 = sheet.getRow(i).getCell(84).getStringCellValue().trim();
	ModelPrice20 = sheet.getRow(i).getCell(85).getStringCellValue().trim();
	UnitSellingPrice20 = sheet.getRow(i).getCell(86).getStringCellValue().trim();

	optionClass3 = sheet.getRow(i).getCell(87).getStringCellValue().trim();

	AddingOptionalItem21 = sheet.getRow(i).getCell(88).getStringCellValue().trim();
	Quantity21 = sheet.getRow(i).getCell(89).getStringCellValue().trim();
	ModelPrice21 = sheet.getRow(i).getCell(90).getStringCellValue().trim();
	UnitSellingPrice21 = sheet.getRow(i).getCell(91).getStringCellValue().trim();
	AddingOptionalItem22 = sheet.getRow(i).getCell(92).getStringCellValue().trim();
	Quantity22 = sheet.getRow(i).getCell(93).getStringCellValue().trim();
	ModelPrice22 = sheet.getRow(i).getCell(94).getStringCellValue().trim();
	UnitSellingPrice22 = sheet.getRow(i).getCell(95).getStringCellValue().trim();
	AddingOptionalItem23 = sheet.getRow(i).getCell(96).getStringCellValue().trim();
	Quantity23 = sheet.getRow(i).getCell(97).getStringCellValue().trim();
	ModelPrice23 = sheet.getRow(i).getCell(98).getStringCellValue().trim();
	UnitSellingPrice23 = sheet.getRow(i).getCell(99).getStringCellValue().trim();
	AddingOptionalItem24 = sheet.getRow(i).getCell(100).getStringCellValue().trim();
	Quantity24 = sheet.getRow(i).getCell(101).getStringCellValue().trim();
	ModelPrice24 = sheet.getRow(i).getCell(102).getStringCellValue().trim();
	UnitSellingPrice24 = sheet.getRow(i).getCell(103).getStringCellValue().trim();
	AddingOptionalItem25 = sheet.getRow(i).getCell(104).getStringCellValue().trim();
	Quantity25 = sheet.getRow(i).getCell(105).getStringCellValue().trim();
	ModelPrice25 = sheet.getRow(i).getCell(106).getStringCellValue().trim();
	UnitSellingPrice25 = sheet.getRow(i).getCell(107).getStringCellValue().trim();
	AddingOptionalItem26 = sheet.getRow(i).getCell(108).getStringCellValue().trim();
	Quantity26 = sheet.getRow(i).getCell(109).getStringCellValue().trim();
	ModelPrice26 = sheet.getRow(i).getCell(110).getStringCellValue().trim();
	UnitSellingPrice26 = sheet.getRow(i).getCell(111).getStringCellValue().trim();
	AddingOptionalItem27 = sheet.getRow(i).getCell(112).getStringCellValue().trim();
	Quantity27 = sheet.getRow(i).getCell(113).getStringCellValue().trim();
	ModelPrice27 = sheet.getRow(i).getCell(114).getStringCellValue().trim();
	UnitSellingPrice27 = sheet.getRow(i).getCell(115).getStringCellValue().trim();
	AddingOptionalItem28 = sheet.getRow(i).getCell(116).getStringCellValue().trim();
	Quantity28 = sheet.getRow(i).getCell(117).getStringCellValue().trim();
	ModelPrice28 = sheet.getRow(i).getCell(118).getStringCellValue().trim();
	UnitSellingPrice28 = sheet.getRow(i).getCell(119).getStringCellValue().trim();
	AddingOptionalItem29 = sheet.getRow(i).getCell(120).getStringCellValue().trim();
	Quantity29 = sheet.getRow(i).getCell(121).getStringCellValue().trim();
	ModelPrice29 = sheet.getRow(i).getCell(122).getStringCellValue().trim();
	UnitSellingPrice29 = sheet.getRow(i).getCell(123).getStringCellValue().trim();
	AddingOptionalItem30 = sheet.getRow(i).getCell(124).getStringCellValue().trim();
	Quantity30 = sheet.getRow(i).getCell(125).getStringCellValue().trim();
	ModelPrice30 = sheet.getRow(i).getCell(126).getStringCellValue().trim();
	UnitSellingPrice30 = sheet.getRow(i).getCell(127).getStringCellValue().trim();
	Bundle_Part_Number = sheet.getRow(i).getCell(132).getStringCellValue();



	LineNumber = sheet.getRow(i).getCell(1).getStringCellValue();
	System.out.println("Salesorder value is :" +SalesOrderItem);
	Addnewitem();
	i++;
	continue;
	}


	}
	 }

	public void Search() throws Exception
	{
	// WebElement search = browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:0:AP1:qq1:it1::content\"]"));
	// highLightElement(browser, search);
	// search.click();
	Thread.sleep(5000);
	WebElement search1 = browser.findElement(By.linkText("Advanced"));
	search1.click();
	WebElement ele = browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:1:AP1:qryId1:operator1::content"));
	Select element = new Select(ele);
	    element.selectByVisibleText("Equals");
	    Thread.sleep(2000);
	    browser.findElement(By.xpath("//input[contains(@id, 'AP1:qryId1:value10::content')]")).sendKeys(Order_Number);
	    browser.findElement(By.xpath("//button[contains(@id, 'AP1:qryId1::search')]")).click();
	    Thread.sleep(5000);
	    browser.findElement(By.linkText(Order_Number)).click();
	    Thread.sleep(8000);
	// WebElement searchicon = browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:0:AP1:qq1::search_icon"));
	// highLightElement(browser, searchicon);
	// searchicon.click();
	// Thread.sleep(5000);
	WebElement actions = browser.findElement(By.xpath("//*[text()='Actions']"));
	highLightElement(browser, actions);
	actions.click();
	WebElement revision = browser.findElement(By.xpath("//*[text()='Create Revision']"));
	highLightElement(browser, revision);
	revision.click();
	Thread.sleep(8000);
	WebElement rev = browser.findElement(By.xpath("//*[contains(@id, 'APRS1:SPph::_afrTtxt')]"));
	String revision1 = rev.getText();
	System.out.println("revision = "+revision1);
	String revisionNum = revision1.substring(revision1.indexOf("(")+1,revision1.indexOf(")"));
	System.out.println("revisionNum = "+revisionNum);
	CellStyle style = wb.createCellStyle();  
	XSSFCell cell = sheet.getRow(i).createCell(131);
	cell.setCellValue(revisionNum);
	Font font = wb.createFont();
	font.setColor(IndexedColors.GREEN.getIndex());
	font.setBold(true);
	style = wb.createCellStyle();
	style.setFont(font);
	cell.setCellStyle(style);
	fos = new FileOutputStream(f);
	wb.write(fos);
	}

	public void cancellItem() throws InterruptedException, IOException
	{
	Thread.sleep(3000);
	JavascriptExecutor js = (JavascriptExecutor) browser;
	WebElement scroll = browser.findElement(By.xpath("//*[contains(@title,'"+SalesOrderItem+"')]/../../../../../../../../../../../../../..//button[contains(@title,'Actions')]"));
	js.executeScript("arguments[0].scrollIntoView();",scroll );
	  scroll.click();
	  Thread.sleep(3000);
	  browser.findElement(By.xpath("(//tr[contains(@id, 'CancelLine')])[1]")).click();
	  try {
	  browser.findElement(By.xpath("//input[contains(@id, 'APRS1:soc4::content')]")).sendKeys(cancelReason);
	  }
	 
	  catch(Exception e)
	  {
	 
	  }
	  browser.findElement(By.xpath("//button[contains(@id, 'APRS1:cb7')]")).click();
	  CellStyle style = wb.createCellStyle();
	XSSFCell cell = sheet.getRow(i).createCell(133);
	cell.setCellValue("PASS");
	Font font = wb.createFont();
	font.setColor(IndexedColors.GREEN.getIndex());
	font.setBold(true);
	style = wb.createCellStyle();
	style.setFont(font);
	cell.setCellStyle(style);
	fos = new FileOutputStream(f);
	wb.write(fos);
	}

	public void Addnewitem() throws Exception
	{
	Thread.sleep(4000);
	browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:searchIcoId::icon")).click();
	Thread.sleep(4000);
	try
	{
	browser.findElement(By.xpath("//*[text()='vanced']")).click();
	Select scdp = new Select(browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp::saveSearch::content\"]")));
	scdp.selectByVisibleText("Application Default copy");
	browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content\"]")).click();
	Thread.sleep(3000);
	browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content\"]")).sendKeys(SalesOrderItem);
	Thread.sleep(4000);
	browser.findElement(By.xpath("//button[text()='Search']")).click();
	}
	catch(Exception e)
	{
	Thread.sleep(4000);
	browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp::_afrDscl")).click();
	browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content\"]")).click();
	browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content\"]")).clear();
	browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp:value00::content\"]")).sendKeys(SalesOrderItem);
	browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:efqrp::search")).click();
	}
	Thread.sleep(5000);
	browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:Popup1:0:Advan1:0:rstab:_ATp:table1::db\"]/table/tbody/tr[1]/td[1]")).click();
	Thread.sleep(4000);
	 browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:itemNumberId:cb1")).click();
	 Thread.sleep(3000);
	 
	 List<String> configAdd = new ArrayList<String>();
	 
	 if(!AddingOptionalItem1.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem1);
	 }
	 if(!AddingOptionalItem2.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem2);
	 }
	 if(!AddingOptionalItem3.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem3);
	 }
	 if(!AddingOptionalItem4.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem4);
	 }
	 if(!AddingOptionalItem5.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem5);
	 }
	 if(!AddingOptionalItem6.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem6);
	 }
	 if(!AddingOptionalItem7.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem7);
	 }
	 if(!AddingOptionalItem8.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem8);
	 }
	 if(!AddingOptionalItem9.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem9);
	 }
	 if(!AddingOptionalItem10.equalsIgnoreCase("NA"))
	 {
	 configAdd.add(AddingOptionalItem10);
	 }
	 
	 List<String> configAdd1 = new ArrayList<String>();
	 
	 if(!AddingOptionalItem11.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem11);
	 }
	 if(!AddingOptionalItem12.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem12);
	 }
	 if(!AddingOptionalItem13.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem13);
	 }
	 if(!AddingOptionalItem14.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem14);
	 }
	 if(!AddingOptionalItem15.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem15);
	 }
	 if(!AddingOptionalItem16.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem16);
	 }
	 if(!AddingOptionalItem17.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem17);
	 }
	 if(!AddingOptionalItem18.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem18);
	 }
	 if(!AddingOptionalItem19.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem19);
	 }
	 if(!AddingOptionalItem20.equalsIgnoreCase("NA"))
	 {
	 configAdd1.add(AddingOptionalItem20);
	 }
	 
	 List<String> configAdd2 = new ArrayList<String>();
	 
	 if(!AddingOptionalItem21.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem21);
	 }
	 if(!AddingOptionalItem22.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem22);
	 }
	 if(!AddingOptionalItem23.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem23);
	 }
	 if(!AddingOptionalItem24.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem24);
	 }
	 if(!AddingOptionalItem25.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem25);
	 }
	 if(!AddingOptionalItem26.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem26);
	 }
	 if(!AddingOptionalItem27.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem27);
	 }
	 if(!AddingOptionalItem28.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem28);
	 }
	 if(!AddingOptionalItem29.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem29);
	 }
	 if(!AddingOptionalItem30.equalsIgnoreCase("NA"))
	 {
	 configAdd2.add(AddingOptionalItem30);
	 }
	 
	 
	 List<String> qnty = new ArrayList<String>();
	 
	 if(!Quantity1.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity1);
	 }
	 if(!Quantity2.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity2);
	 }
	 if(!Quantity3.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity3);
	 }
	 if(!Quantity4.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity4);
	 }
	 if(!Quantity5.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity5);
	 }
	 if(!Quantity6.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity6);
	 }
	 if(!Quantity7.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity7);
	 }
	 if(!Quantity8.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity8);
	 }
	 if(!Quantity9.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity9);
	 }
	 if(!Quantity10.equalsIgnoreCase("NA"))
	 {
	 qnty.add(Quantity10);
	 }
	 
	 List<String> qnty1 = new ArrayList<String>();
	 
	 if(!Quantity11.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity11);
	 }
	 if(!Quantity12.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity12);
	 }
	 if(!Quantity13.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity13);
	 }
	 if(!Quantity14.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity14);
	 }
	 if(!Quantity15.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity15);
	 }
	 if(!Quantity16.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity16);
	 }
	 if(!Quantity17.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity17);
	 }
	 if(!Quantity18.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity18);
	 }
	 if(!Quantity19.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity19);
	 }
	 if(!Quantity20.equalsIgnoreCase("NA"))
	 {
	 qnty1.add(Quantity20);
	 }
	 
	 List<String> qnty2 = new ArrayList<String>();
	 
	 if(!Quantity21.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity21);
	 }
	 if(!Quantity22.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity22);
	 }
	 if(!Quantity23.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity23);
	 }
	 if(!Quantity24.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity24);
	 }
	 if(!Quantity25.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity25);
	 }
	 if(!Quantity26.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity26);
	 }
	 if(!Quantity27.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity27);
	 }
	 if(!Quantity28.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity28);
	 }
	 if(!Quantity29.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity29);
	 }
	 if(!Quantity30.equalsIgnoreCase("NA"))
	 {
	 qnty2.add(Quantity30);
	 }
	 
	 List<String> modelPrice = new ArrayList<String>();
	 
	 if(!ModelPrice1.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice1);
	 }
	 if(!ModelPrice2.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice2);
	 }
	 if(!ModelPrice3.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice3);
	 }
	 if(!ModelPrice4.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice4);
	 }
	 if(!ModelPrice5.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice5);
	 }
	 if(!ModelPrice6.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice6);
	 }
	 if(!ModelPrice7.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice7);
	 }
	 if(!ModelPrice8.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice8);
	 }
	 if(!ModelPrice9.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice9);
	 }
	 if(!ModelPrice10.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice10);
	 }
	 if(!ModelPrice11.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice11);
	 }
	 if(!ModelPrice12.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice12);
	 }
	 if(!ModelPrice13.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice13);
	 }
	 if(!ModelPrice14.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice14);
	 }
	 if(!ModelPrice15.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice15);
	 }
	 if(!ModelPrice16.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice16);
	 }
	 if(!ModelPrice17.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice17);
	 }
	 if(!ModelPrice18.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice18);
	 }
	 if(!ModelPrice19.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice19);
	 }
	 if(!ModelPrice20.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice20);
	 }
	 if(!ModelPrice21.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice21);
	 }
	 if(!ModelPrice22.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice22);
	 }
	 if(!ModelPrice23.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice23);
	 }
	 if(!ModelPrice24.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice24);
	 }
	 if(!ModelPrice25.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice25);
	 }
	 if(!ModelPrice26.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice26);
	 }
	 if(!ModelPrice27.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice27);
	 }
	 if(!ModelPrice28.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice28);
	 }
	 if(!ModelPrice29.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice29);
	 }
	 if(!ModelPrice30.equalsIgnoreCase("NA"))
	 {
	 modelPrice.add(ModelPrice30);
	 }
	 
	        List<String> unitSelling = new ArrayList<String>();
	 
	 if(!UnitSellingPrice1.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice1);
	 }
	 if(!UnitSellingPrice2.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice2);
	 }
	 if(!UnitSellingPrice3.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice3);
	 }
	 if(!UnitSellingPrice4.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice4);
	 }
	 if(!UnitSellingPrice5.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice5);
	 }
	 if(!UnitSellingPrice6.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice6);
	 }
	 if(!UnitSellingPrice7.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice7);
	 }
	 if(!UnitSellingPrice8.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice8);
	 }
	 if(!UnitSellingPrice9.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice9);
	 }
	 if(!UnitSellingPrice10.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice10);
	 }
	 if(!UnitSellingPrice11.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice11);
	 }
	 if(!UnitSellingPrice12.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice12);
	 }
	 if(!UnitSellingPrice13.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice13);
	 }
	 if(!UnitSellingPrice14.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice14);
	 }
	 if(!UnitSellingPrice15.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice15);
	 }
	 if(!UnitSellingPrice16.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice16);
	 }
	 if(!UnitSellingPrice17.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice17);
	 }
	 if(!UnitSellingPrice18.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice18);
	 }
	 if(!UnitSellingPrice19.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice19);
	 }
	 if(!UnitSellingPrice20.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice20);
	 }
	 if(!UnitSellingPrice21.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice21);
	 }
	 if(!UnitSellingPrice22.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice22);
	 }
	 if(!UnitSellingPrice23.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice23);
	 }
	 if(!UnitSellingPrice24.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice24);
	 }
	 if(!UnitSellingPrice25.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice25);
	 }
	 if(!UnitSellingPrice26.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice26);
	 }
	 if(!UnitSellingPrice27.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice27);
	 }
	 if(!UnitSellingPrice28.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice28);
	 }
	 if(!UnitSellingPrice29.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice29);
	 }
	 if(!UnitSellingPrice30.equalsIgnoreCase("NA"))
	 {
	 unitSelling.add(UnitSellingPrice30);
	 }
	 
	 
	 //Code for ATO model
	 
	 try
	{
	 Thread.sleep(3000);
	 WebElement qny=browser.findElement(By.xpath("//input[contains(@id, 'APRS1:createLineQuantity::content')]"));
	 qny.click();
	 qny.clear();
	//  Thread.sleep(3000);
	 qny.sendKeys(ATOModel_Qnty);
	 Thread.sleep(5000);
	WebElement configAddbutton = browser.findElement(By.xpath("//span[text()='Configure and Add']"));
	JavascriptExecutor js=(JavascriptExecutor)browser;
	js.executeScript("arguments[0].scrollIntoView();",configAddbutton );
	configAddbutton.click();
	Thread.sleep(7000);
	for(int k=0; k<configAdd.size(); k++)
	{
	try {
	Thread.sleep(3000);

	if(!optionClass1.equalsIgnoreCase("NA"))
	{
	try {
	if(k==0)
	{
	WebElement option1 = browser.findElement(By.xpath("//*[contains(text(), '"+optionClass1+"')]/../../../../../../..//img[contains(@title,'Query By Example')]"));
	JavascriptExecutor jsm=(JavascriptExecutor)browser;
	jsm.executeScript("arguments[0].scrollIntoView();",option1 );
	option1.click();
	}
	Thread.sleep(4000);
	WebElement filter = browser.findElement(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	filter.click();
	filter.clear();
	Thread.sleep(3000);
	filter.sendKeys(configAdd.get(k));
	filter.sendKeys(Keys.ENTER);
	Thread.sleep(5000);
	try {
	WebElement checkbox = browser.findElement(By.xpath("//*[contains(text(), '"+configAdd.get(k)+"')]/../../../..//img[contains(@title,'Select')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd.get(k)+"')]/../../../../../..//input"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty.get(k));
	}
	catch(Exception e)
	{
	WebElement checkbox = browser.findElement(By.xpath("//*[contains(text(), '"+configAdd.get(k)+"')]/../../../..//*[contains(@type,'checkbox')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd.get(k)+"')]/../../../../../..//input[contains(@type,'text')]"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty.get(k));
	}

	}

	catch(Exception exe)
	{
	WebElement cnfitem = browser.findElement(By.xpath("//td[text()='"+configAdd.get(k)+"']"));
	JavascriptExecutor js1=(JavascriptExecutor)browser;
	js1.executeScript("arguments[0].scrollIntoView();",cnfitem );
	Thread.sleep(3000);
	try {
	WebElement checkbox = browser.findElement(By.xpath("//td[text()='"+configAdd.get(k)+"']/..//img[contains(@title,'Select')]"));
	JavascriptExecutor js3=(JavascriptExecutor)browser;
	js3.executeScript("arguments[0].scrollIntoView();",checkbox );
	js3.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd.get(k)+"')]/..//input"));
	js3.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty.get(k));
	}
	catch(Exception e)
	{
	WebElement checkbox = browser.findElement(By.xpath("//td[text()='"+configAdd.get(k)+"']/..//*[contains(@type,'checkbox')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd.get(k)+"')]/..//input[contains(@type,'text')]"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty.get(k));
	}
	}
	}
	}
	catch(Exception e)
	{

	}
	}
	try {
	List<WebElement>ele = browser.findElements(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	if(ele.size()>0)
	{
	WebElement filter = browser.findElement(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	filter.click();
	filter.clear();
	filter.sendKeys(Keys.ENTER);
	Thread.sleep(3000);
	WebElement option1 = browser.findElement(By.xpath("//*[contains(text(), '"+optionClass1+"')]/../../../../../../..//img[contains(@title,'Query By Example')]"));
	JavascriptExecutor jsm1=(JavascriptExecutor)browser;
	jsm1.executeScript("arguments[0].scrollIntoView();",option1 );
	option1.click();
	}

	}
	catch(Exception e)
	{

	}

	for(int k=0; k<configAdd1.size(); k++)
	{
	try {
	Thread.sleep(3000);

	if(!optionClass2.equalsIgnoreCase("NA"))
	{
	try {
	if(k==0)
	{
	WebElement option2 = browser.findElement(By.xpath("//*[contains(text(), '"+optionClass2+"')]/../../../../../../..//img[contains(@title,'Query By Example')]"));
	JavascriptExecutor jsm1=(JavascriptExecutor)browser;
	jsm1.executeScript("arguments[0].scrollIntoView();",option2 );
	option2.click();
	}
	   Thread.sleep(4000);
	WebElement filter = browser.findElement(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	filter.click();
	filter.clear();
	// Thread.sleep(3000);
	filter.sendKeys(configAdd1.get(k));
	filter.sendKeys(Keys.ENTER);
	Thread.sleep(5000);
	try {
	WebElement checkbox = browser.findElement(By.xpath("//*[contains(text(), '"+configAdd1.get(k)+"')]/../../../..//img[contains(@title,'Select')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd1.get(k)+"')]/../../../../../..//input"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty1.get(k));
	}
	catch(Exception e)
	{
	WebElement checkbox = browser.findElement(By.xpath("//*[contains(text(), '"+configAdd1.get(k)+"')]/../../../..//*[contains(@type,'checkbox')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd1.get(k)+"')]/../../../../../..//input[contains(@type,'text')]"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty1.get(k));
	}

	}
	catch(Exception exe)
	{
	WebElement cnfitem = browser.findElement(By.xpath("//td[text()='"+configAdd1.get(k)+"']"));
	JavascriptExecutor js1=(JavascriptExecutor)browser;
	js1.executeScript("arguments[0].scrollIntoView();",cnfitem );
	Thread.sleep(3000);
	try {
	WebElement checkbox = browser.findElement(By.xpath("//td[text()='"+configAdd1.get(k)+"']/..//img[contains(@title,'Select')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd1.get(k)+"')]/..//input"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty1.get(k));
	}
	catch(Exception e)
	{
	WebElement checkbox = browser.findElement(By.xpath("//td[text()='"+configAdd1.get(k)+"']/..//*[contains(@type,'checkbox')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd1.get(k)+"')]/..//input[contains(@type,'text')]"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty1.get(k));
	}
	}
	}
	}
	catch(Exception e)
	{

	}
	}
	try {
	List<WebElement>ele = browser.findElements(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	if(ele.size()>0)
	{
	WebElement filter = browser.findElement(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	filter.click();
	filter.clear();
	filter.sendKeys(Keys.ENTER);
	Thread.sleep(3000);
	WebElement option2 = browser.findElement(By.xpath("//*[contains(text(), '"+optionClass2+"')]/../../../../../../..//img[contains(@title,'Query By Example')]"));
	JavascriptExecutor jsm1=(JavascriptExecutor)browser;
	jsm1.executeScript("arguments[0].scrollIntoView();",option2 );
	option2.click();
	}

	}
	catch(Exception e)
	{

	}

	for(int k=0; k<configAdd2.size(); k++)
	{
	try {
	Thread.sleep(3000);

	if(!optionClass3.equalsIgnoreCase("NA"))
	{
	try {
	if(k==0)
	{
	WebElement option3 = browser.findElement(By.xpath("//*[contains(text(), '"+optionClass3+"')]/../../../../../../..//img[contains(@title,'Query By Example')]"));
	JavascriptExecutor jsm1=(JavascriptExecutor)browser;
	jsm1.executeScript("arguments[0].scrollIntoView();",option3 );
	option3.click();
	}
	   Thread.sleep(4000);
	WebElement filter = browser.findElement(By.xpath("//input[contains(@id, 'AT1:_ATp:ist:dc_it1::content')]"));
	filter.click();
	filter.clear();
	// Thread.sleep(3000);
	filter.sendKeys(configAdd2.get(k));
	filter.sendKeys(Keys.ENTER);
	Thread.sleep(5000);
	try {
	WebElement checkbox = browser.findElement(By.xpath("//*[contains(text(), '"+configAdd2.get(k)+"')]/../../../..//img[contains(@title,'Select')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd2.get(k)+"')]/../../../../../..//input"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty2.get(k));
	}
	catch(Exception e)
	{
	WebElement checkbox = browser.findElement(By.xpath("//*[contains(text(), '"+configAdd2.get(k)+"')]/../../../..//*[contains(@type,'checkbox')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd2.get(k)+"')]/../../../../../..//input[contains(@type,'text')]"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(2000);
	edit.sendKeys(qnty2.get(k));
	}

	}
	catch(Exception exe)
	{
	WebElement cnfitem = browser.findElement(By.xpath("//td[text()='"+configAdd2.get(k)+"']"));
	JavascriptExecutor js1=(JavascriptExecutor)browser;
	js1.executeScript("arguments[0].scrollIntoView();",cnfitem );
	Thread.sleep(3000);
	try {
	WebElement checkbox = browser.findElement(By.xpath("//td[text()='"+configAdd2.get(k)+"']/..//img[contains(@title,'Select')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd2.get(k)+"')]/..//input"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty2.get(k));
	}
	catch(Exception e)
	{
	WebElement checkbox = browser.findElement(By.xpath("//td[text()='"+configAdd2.get(k)+"']/..//*[contains(@type,'checkbox')]"));
	JavascriptExecutor js2=(JavascriptExecutor)browser;
	js2.executeScript("arguments[0].scrollIntoView();",checkbox );
	js2.executeScript("arguments[0].click()", checkbox);
	// checkbox.click();
	Thread.sleep(5000);
	WebElement edit = browser.findElement(By.xpath("//td[contains(text(),'"+configAdd2.get(k)+"')]/..//input[contains(@type,'text')]"));
	js2.executeScript("arguments[0].click()", edit);
	// edit.click();
	edit.clear();
	Thread.sleep(4000);
	edit.sendKeys(qnty2.get(k));
	}
	}
	}
	}
	catch(Exception e)
	{

	}
	}


	Thread.sleep(8000);
	JavascriptExecutor js2 = (JavascriptExecutor)browser;
	js2.executeScript("window.scrollBy(0,-1000)");
	browser.findElement(By.xpath("//span[text()='Finish']")).click();
	Thread.sleep(8000);
	List<WebElement> tablerows1 = browser.findElements(By.xpath("//*[contains(@id, 'APRS1:pc1:t1::db')]/table/tbody/tr"));
	tabler = tablerows1.size();
	System.out.println("Size of table :" +tabler);
	System.out.println("Item value is :" +SalesOrderItem);
	WebElement ele2 = browser.findElement(By.xpath("//span[contains(text(), 'Changed')]/../..//span[contains(text(),'"+SalesOrderItem+"')]/../../../../../../../../../../..//a[contains(text(),'More...')]"));
	Thread.sleep(5000);
	if(ele2.getText().contains("More...")) {
		try {
			browser.findElement(By.xpath("(//span[contains(text(), 'Changed')]/../..//span[contains(text(),'"+SalesOrderItem+"')]/../../../../../../../../../../..//a[contains(text(),'More...')])[2]")).click();
		}
		catch(Exception e)
		{
			browser.findElement(By.xpath("//span[contains(text(), 'Changed')]/../..//span[contains(text(),'"+SalesOrderItem+"')]/../../../../../../../../../../..//a[contains(text(),'More...')]")).click();
		}
	
	Thread.sleep(5000);
	List<WebElement> tablerows = browser.findElements(By.xpath("//*[contains(@id, 'AP1:pc1:dc_tt1::db')]/table/tbody/tr"));
	System.out.println("Rows size is :" +tablerows.size());
	int rowvalue = tablerows.size();
	
	for(int n =2; n<=rowvalue; n++)
	{
		WebElement ele = browser.findElement(By.xpath("//*[contains(@id, 'AP1:pc1:dc_tt1::db')]/table/tbody/tr["+ n +"]/td[1]")); 
		String value = ele.getText().trim();
		System.out.println("Table value :" +value);
		Thread.sleep(3000);
		JavascriptExecutor jsp = (JavascriptExecutor)browser;
		jsp.executeScript("arguments[0].scrollIntoView();", ele);
		jsp.executeScript("arguments[0].click()", ele);
		WebElement button1 = browser.findElement(By.xpath("//span[text()='"+value+"']/../../../../..//button[contains(@title,'Actions')]"));
		JavascriptExecutor jsprice1 = (JavascriptExecutor)browser;
		jsprice1.executeScript("arguments[0].scrollIntoView();", button1);
		button1.click();
		Thread.sleep(4000);
		browser.findElement(By.xpath("//tr[contains(@id, 'AP1:pc1:dc_tt1:"+(n-1)+":flineAdditionalInfo')]")).click();
	    Thread.sleep(4000);
	    browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).click();
	    browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).sendKeys(Bundle_Part_Number);
	    Thread.sleep(3000);
	    browser.findElement(By.xpath("//button[contains(@id, 'dEffAttr::ok')]")).click();
	    Thread.sleep(3000);
	}
	int s=0;
	System.out.println("size=="+modelPrice.size());

	for(int m =1; m<=rowvalue; m++)
	{
	try {
	WebElement ele = browser.findElement(By.xpath("//*[contains(@id, 'AP1:pc1:dc_tt1::db')]/table/tbody/tr["+ m +"]/td[1]"));                             
	String value = ele.getText().trim();
	System.out.println("Table value :" +value);

	if(modelPrice.get(s).equalsIgnoreCase(value) && s<modelPrice.size())
	 {
	Thread.sleep(5000);
	JavascriptExecutor jsp = (JavascriptExecutor)browser;
	jsp.executeScript("arguments[0].scrollIntoView();", ele);
	jsp.executeScript("arguments[0].click()", ele);
	// ele.click();
	WebElement button1 = browser.findElement(By.xpath("//span[text()='"+value+"']/../../../../..//button[contains(@title,'Actions')]"));
	JavascriptExecutor jsprice1 = (JavascriptExecutor)browser;
	jsprice1.executeScript("arguments[0].scrollIntoView();", button1);
	button1.click();
	Thread.sleep(5000);
	browser.findElement(By.xpath("//tr[contains(@id, 'AP1:pc1:dc_tt1:"+(m-1)+":flineAdditionalInfo')]")).click();
	   Thread.sleep(5000);
	   browser.findElement(By.linkText("Pricing Additional Information")).click();
	   Thread.sleep(3000);
	WebElement Price = browser.findElement(By.xpath("//input[contains(@id, 'unitsellingprice')]"));                        
	jsp.executeScript("arguments[0].click()", Price);
	   Price.clear();
	   Thread.sleep(3000);
	   Price.sendKeys(unitSelling.get(s));
	   browser.findElement(By.xpath("//button[contains(@id, 'AP1:dEffAttr::ok')]")).click();
	s++;
	m=1;
	 }

	 }
	catch (IndexOutOfBoundsException error)
	{

	}
	catch(Exception e)
	{
	 e.printStackTrace();
	}}

	Thread.sleep(8000);
	browser.findElement(By.xpath("//*[contains(@id, 'AP1:dc_ctb1')]/a/span")).click();
	// i++;
	}
	}
	 
	 //Code for Standard item
	 
	 catch(Exception e)
	 {
		 
		 Thread.sleep(4000);
//		 browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:createLineQuantity::content\"]")).click();
//		 browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:createLineQuantity::content\"]")).clear();
//		 WebElement qty = browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:createLineQuantity::content\"]"));
//		 qty.sendKeys(Quantity1);
		  browser.findElement(By.xpath("//span[text()='Add']")).click();
		   Thread.sleep(6000);
		 List<WebElement> tablerows2 = browser.findElements(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:pc1:t1::db\"]/table/tbody/tr"));
		 Thread.sleep(3000);
		 tabler = tablerows2.size();
		 System.out.println("Size of table :" +tabler);
		 Thread.sleep(3000);
		   browser.findElement(By.xpath("//*[@id=\"pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:pc1:t1::db\"]/table/tbody/tr["+ (tabler) +"]/td[1]")).click();
		   try
		   {
               WebElement ele3 = browser.findElement(By.xpath("(//img[contains(@title,'NMX-HWP-3G-A-03')]/../../../../../../../../..//button[contains(@title,'Actions')])[2]"));
               JavascriptExecutor jse = (JavascriptExecutor)browser;
      		   jse.executeScript("arguments[0].scrollIntoView()", ele3);
      		   ele3.click();
		   }
		   catch(Exception e1)
		   {
			     WebElement ele = browser.findElement(By.xpath("//img[contains(@title,'"+SalesOrderItem+"')]/../../../../../../../../..//button[contains(@title,'Actions')]"));
				 JavascriptExecutor jse = (JavascriptExecutor)browser;
				 jse.executeScript("arguments[0].scrollIntoView()", ele);
				 ele.click();
		   }
		 Thread.sleep(3000);
		 browser.findElement(By.xpath("//tr[contains(@id,'lineAdditionalInfo')]")).click();
		 Thread.sleep(3000);
		 browser.findElement(By.linkText("Pricing Additional Information")).click();
		 Thread.sleep(5000);
		 WebElement elpr = browser.findElement(By.xpath("//input[contains(@id, 'unitsellingprice')]"));
		 elpr.click();
		 elpr.clear();
		 elpr.sendKeys(UnitSellingPrice1);
		 Thread.sleep(3000);
		 browser.findElement(By.linkText("Global Data Elements")).click();
		 Thread.sleep(2000);
		 browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).click();
		 browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).clear();
		 browser.findElement(By.xpath("//*[contains(@id,'CustomerPartNumber::content')]")).sendKeys(Bundle_Part_Number);
		 Thread.sleep(3000);
		 browser.findElement(By.id("pt1:_FOr1:1:_FOSritemNode_order_management_order_management:0:_FOTsr1:3:APRS1:dEffAttr::ok")).click();
		 JavascriptExecutor jsvertical = (JavascriptExecutor)browser;
		 jsvertical.executeScript("window.scrollBy(-1000,0)","");
//	 e.printStackTrace();
	  }
	 CellStyle style = wb.createCellStyle();
	 XSSFCell cell = sheet.getRow(i).createCell(133);
	cell.setCellValue("PASS");
	Font font = wb.createFont();
	font.setColor(IndexedColors.GREEN.getIndex());
	font.setBold(true);
	style = wb.createCellStyle();
	style.setFont(font);
	cell.setCellStyle(style);
	fos = new FileOutputStream(f);
	wb.write(fos);
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
	@AfterTest()
	public void Close_Browser()
	{
	// browser.quit();
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
