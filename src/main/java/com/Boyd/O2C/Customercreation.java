package com.Boyd.O2C;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.TimeoutException;
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
import org.testng.Assert;

@SuppressWarnings("unused")
public class Customercreation {
	public WebDriver browser;
	public String Name;
	public String Account_Description;
	public String Account_Type;
	public String Account_Address_Set;
	public String Label_Format;
	public String Address_Line1;
	public String City;
	public String State;
	public String Postal_Code;
	public String Purpose;
	public String Purpose1;
	public String First_Name;
	public String Last_Name;
	public String Contact_Point_Type;
	public String Email;

	public static WebDriverWait wait;
	public static int timeout = 60;
	public static int lag = 3;
	public File f;
	public XSSFWorkbook wb;
	public XSSFSheet sheet;
     int i;

	@BeforeTest()
	public void Login_Page() throws Exception {
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(100, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(100, TimeUnit.SECONDS);
		
		browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
//		String url = System.getProperty("url");
//		System.out.println("url is :" +url);
//		String username = System.getProperty("userName");
//		System.out.println("username  is :" +username);
//		String password = System.getProperty("password");
//		System.out.println("password  is :" +password);
		
//		browser.get(url);
		browser.findElement(By.id("userid")).click();	
		//browser.findElement(By.id("userid")).sendKeys("forsys.user");
		
		
//		browser.findElement(By.id("userid")).sendKeys(username);
//		browser.findElement(By.id("password")).click();
//		//browser.findElement(By.id("password")).sendKeys("welcome123");	
//		browser.findElement(By.id("password")).sendKeys(password);
//		
//		browser.findElement(By.id("btnActive")).click();
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("welcome123");
		browser.findElement(By.id("btnActive")).click();
		WebElement homepage = browser.findElement(By.xpath("//a[text()='You have a new home page!']"));
		waitUntilElementClickable("homepage", homepage, browser, timeout);
		WebElement receivable = browser.findElement(By.linkText("Receivables"));
		waitUntilElementClickable("receivable", receivable, browser, timeout);
		WebElement bill = browser.findElement(By.linkText("Billing"));
		waitUntilElementClickable("bill", bill, browser, timeout);
	}

	@Test()
	public void Home_Page() throws Exception {
		try{
		
		
//		String remoteExecution = System.getProperty("remoteExecution");
//		System.out.println("remoteFlag  is :" +remoteExecution);
//		Boolean exectionFlag = Boolean.parseBoolean(remoteExecution);
//		
//		if(exectionFlag) {
//			System.out.println("remote execution");
//			f = new File(System.getProperty("user.dir") + "\\ExternalFiles\\BOYD_P2P_Customer_Creation.xlsx");
//		}			 
//		else {
//			 f = new File(System.getProperty("user.dir") + "\\Excel\\BOYD_P2P_Customer_Creation.xlsx");
//			 System.out.println("local execution");
//		}
//		FileInputStream fis = new FileInputStream(f);
//		wb = new XSSFWorkbook(fis);
//		sheet = wb.getSheet("Customer_Creation");
//		sheet.getRow(0).createCell(15).setCellValue("Result");
//		int totalrows = sheet.getPhysicalNumberOfRows();
//		System.out.println("Total number of Excel rows are :" + totalrows);
			File f = new File(System.getProperty("user.dir") + "\\Excel\\BOYD_P2P_Customer_Creation.xlsx");
			FileInputStream fis = new FileInputStream(f);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet("Customer_Creation");
			sheet.getRow(0).createCell(15).setCellValue("Result");
			int totalrows = sheet.getPhysicalNumberOfRows();
			System.out.println("Total number of Excel rows are :" + totalrows);
		if (sheet.getRow(1).getCell(15) == null) {

			for (i = 1; i <= totalrows; i++) {

				if (sheet.getRow(i) == null) {
					return;
				}

				Name = sheet.getRow(i).getCell(0).getStringCellValue();
				Account_Description = sheet.getRow(i).getCell(1).getStringCellValue();
				Account_Type = sheet.getRow(i).getCell(2).getStringCellValue();
				Account_Address_Set = sheet.getRow(i).getCell(3).getStringCellValue();
				Label_Format = sheet.getRow(i).getCell(4).getStringCellValue();
				Address_Line1 = sheet.getRow(i).getCell(5).getStringCellValue();
				City = sheet.getRow(i).getCell(6).getStringCellValue();
				State = sheet.getRow(i).getCell(7).getStringCellValue();
				Postal_Code = sheet.getRow(i).getCell(8).getStringCellValue();
				Purpose = sheet.getRow(i).getCell(9).getStringCellValue();
				Purpose1 = sheet.getRow(i).getCell(10).getStringCellValue();
				First_Name = sheet.getRow(i).getCell(11).getStringCellValue();
				Last_Name = sheet.getRow(i).getCell(12).getStringCellValue();
				Contact_Point_Type = sheet.getRow(i).getCell(13).getStringCellValue();
				Email = sheet.getRow(i).getCell(14).getStringCellValue();

				WebElement task = browser.findElement(
						By.xpath("//*[contains(@id,'_FOTsdi__TransactionsWorkArea_itemNode__FndTasksList::icon')]"));
				waitUntilElementClickable("task", task, browser, timeout);
				WebElement customer = browser.findElement(By.linkText("Create Customer"));
				waitUntilElementClickable("customer", customer, browser, timeout);
				WebElement name = browser.findElement(By.xpath("(//*[contains(@id,'inputText123::content')])[1]"));
				waitUntilElementClickable("name", name, browser, timeout);
				WaituntilElementwritable("name", name, browser, Name);
				WebElement account = browser.findElement(By.xpath("(//*[contains(@id,'inputText2::content')])[1]"));
				waitUntilElementClickable("account", account, browser, timeout);
				WaituntilElementwritable("account", account, browser, Account_Description);
				Select accounttype = new Select(
						browser.findElement(By.xpath("(//*[contains(@id,'selectOneChoice1::content')])[1]")));
				accounttype.selectByVisibleText(Account_Type);
				WebElement searchdropdown = browser.findElement(By.xpath("//*[contains(@id,'setIdLovId::lovIconId')]"));
				waitUntilElementClickable("searchdropdown", searchdropdown, browser, timeout);
				WebElement Search = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("Search", Search, browser, timeout);
				WebElement setname = browser.findElement(
						By.xpath("//*[contains(@id,'setIdLovId::_afrLovInternalQueryId:value10::content')]"));
				waitUntilElementClickable("setname", setname, browser, timeout);
				WaituntilElementwritable("setname", setname, browser, Account_Address_Set);
				WebElement dropdownsearch = browser
						.findElement(By.xpath("//*[contains(@id,'setIdLovId::_afrLovInternalQueryId::search')]"));
				waitUntilElementClickable("dropdownsearch", dropdownsearch, browser, timeout);
				WebElement tablerow = browser.findElement(
						By.xpath("//*[contains(@id,'setIdLovId_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("tablerow", tablerow, browser, timeout);
				WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'setIdLovId::lovDialogId::ok')]"));
				waitUntilElementClickable("okbutton", okbutton, browser, timeout);
				WebElement icon = browser.findElement(By.xpath(
						"//*[contains(@id,'hzdf40_CustAcctSiteInformationIteratorlabelFormat__FLEX_EMPTY::lovIconId')]"));
				waitUntilElementClickable("icon", icon, browser, timeout);
				WebElement labelsearch = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("labelsearch", labelsearch, browser, timeout);
				WebElement value = browser.findElement(By.xpath(
						"//*[contains(@id,'hzdf40_CustAcctSiteInformationIteratorlabelFormat__FLEX_EMPTY::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("value", value, browser, timeout);
				WaituntilElementwritable("value", value, browser, Label_Format);
				WebElement searchicon = browser.findElement(By.xpath(
						"//*[contains(@id,'hzdf40_CustAcctSiteInformationIteratorlabelFormat__FLEX_EMPTY::_afrLovInternalQueryId::search')]"));
				waitUntilElementClickable("searchicon", searchicon, browser, timeout);
				WebElement valuetable = browser.findElement(By.xpath(
						"//*[contains(@id,'hzdf40_CustAcctSiteInformationIteratorlabelFormat__FLEX_EMPTY_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				waitUntilElementClickable("valuetable", valuetable, browser, timeout);
				WebElement button = browser.findElement(By.xpath(
						"//*[contains(@id,'hzdf40_CustAcctSiteInformationIteratorlabelFormat__FLEX_EMPTY::lovDialogId::ok')]"));
				waitUntilElementClickable("button", button, browser, timeout);
				WebElement addressline = browser.findElement(By.xpath("(//*[contains(@id,'inputText2::content')])[2]"));
				waitUntilElementClickable("addressline", addressline, browser, timeout);
				WaituntilElementwritable("addressline", addressline, browser, Address_Line1);
				
				WebElement cityiconlabel = browser.findElement(By.xpath("//label[text()='City']//parent::td//following-sibling::td//a[contains(@id,'lovIconId')]"));				
				waitUntilElementClickable("cityvalue", cityiconlabel, browser, timeout);
				//WaituntilElementwritable("cityvalue", cityvalue, browser, City);
				
				WebElement citysearch = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("statesearch", citysearch, browser, timeout);
				
				WebElement cityvalue = browser.findElement(By.xpath(
						"//input[contains(@id,'afrLovInternalQueryId:value')]"));
				waitUntilElementClickable("statevalue", cityvalue, browser, timeout);
				cityvalue.clear();
				WaituntilElementwritable("statevalue", cityvalue, browser, City);
				WebElement searchcityvalue = browser.findElement(
						By.xpath("//button[contains(@id,'afrLovInternalQueryId::search')]"));
				waitUntilElementClickable("searchvalue", searchcityvalue, browser, timeout);
				
//				By cityRow = By.xpath("//span[contains(text(),'" + Account_Description + "')]");
//				WaituntilElementIsDisplayed("tablerowCity_Displayed", cityRow, browser);
				WebElement tablevalue = browser.findElement(By.xpath("//*[contains(@id,'inputComboboxListOfValues3_afrLovInternalTableId::db')]/table/tbody/tr[1]/td[1]"));
				waitUntilElementClickable("tablevalue", tablevalue, browser, timeout);
				
				WebElement okvalue = browser
						.findElement(By.xpath("//button[contains(@id,'lovDialogId::ok')]"));
				waitUntilElementClickable("okvalue", okvalue, browser, timeout);
				
				By byOK = By.xpath("//button[contains(@id,'lovDialogId::ok')]");
				waitUntilElementDisappearCount("OKButton_Disapper", byOK, browser);
				
				//Thread.sleep(3000);
				//browser.findElement(By.xpath("//*[contains(@id,'inputComboboxListOfValues3::content')]")).sendKeys(Keys.ENTER);
				
				
				
				WebElement iconlabel = browser
						.findElement(By.xpath("//*[contains(@id,'inputComboboxListOfValues1::lovIconId')]"));
				waitUntilElementClickable("iconlabel", iconlabel, browser, timeout);
				WebElement statesearch = browser.findElement(By.linkText("Search..."));
				waitUntilElementClickable("statesearch", statesearch, browser, timeout);
				WebElement statevalue = browser.findElement(By.xpath(
						"//*[contains(@id,'inputComboboxListOfValues1::_afrLovInternalQueryId:value00::content')]"));
				waitUntilElementClickable("statevalue", statevalue, browser, timeout);
				statevalue.clear();
				WaituntilElementwritable("statevalue", statevalue, browser, State);
				WebElement searchvalue = browser.findElement(
						By.xpath("//*[contains(@id,'inputComboboxListOfValues1::_afrLovInternalQueryId::search')]"));
				waitUntilElementClickable("searchvalue", searchvalue, browser, timeout);
				
//			WebElement tablerowvalue = browser.findElement(By.xpath("//*[contains(@id,'inputComboboxListOfValues1_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				By tablerowvalue1 = By.xpath("//span[text()='" + State + "']");
				WaituntilElementIsDisplayed("tablerowvalue1_Displayed", tablerowvalue1, browser);

				okvalue = browser
						.findElement(By.xpath("//*[contains(@id,'inputComboboxListOfValues1::lovDialogId::ok')]"));
				waitUntilElementClickable("okvalue", okvalue, browser, timeout);

				byOK = By.xpath("//*[contains(@id,'inputComboboxListOfValues1::lovDialogId::ok')]");
				waitUntilElementDisappearCount("OKButton_Disapper", byOK, browser);
				
				
				
				// Thread.sleep(10000);
				WebElement postalcode = browser
						.findElement(By.xpath("//*[contains(@id,'inputComboboxListOfValues4::content')]"));

				waitUntilElementClickable("postalcode", postalcode, browser, timeout);
				WaituntilElementwritable("postalcode", postalcode, browser, Postal_Code);
				WebElement createicon = browser.findElement(By.xpath("//*[contains(@id,'AT1:_ATp:create::icon')]"));
				waitUntilElementClickable("createicon", createicon, browser, timeout);
				Select sc = new Select(browser.findElement(By.xpath("//*[contains(@id,'SiteUseCode::content')]")));
				sc.selectByVisibleText(Purpose);
				WebElement iconvalue = browser.findElement(By.xpath("//*[contains(@id,'AT1:_ATp:create::icon')]"));
				waitUntilElementClickable("iconvalue", iconvalue, browser, timeout);
				Select sc1 = new Select(
						browser.findElement(By.xpath("//*[contains(@id,'table1:1:SiteUseCode::content')]")));
				sc1.selectByVisibleText(Purpose1);
				JavascriptExecutor js = (JavascriptExecutor) browser;
				//js.executeScript("window.scrollBy(0,-450)");
				
				WebElement saveclose = browser.findElement(By.xpath("//*[text()='ave and Close']"));
				js.executeScript("arguments[0].scrollIntoView(true);", saveclose);
				waitUntilElementClickable("saveclose", saveclose, browser, timeout);
				JavascriptExecutor down = (JavascriptExecutor) browser;
				//down.executeScript("window.scrollBy(0,350)");
			
				WebElement outputtext = browser.findElement(By.xpath("//*[contains(@id,'outputText61')]"));
				
				down.executeScript("arguments[0].scrollIntoView(true);", outputtext);
				waitUntilElementClickable("outputtext", outputtext, browser, timeout);
				WebElement communication = browser.findElement(By.linkText("Communication"));
				waitUntilElementClickable("communication", communication, browser, timeout);
				WebElement edit = browser.findElement(By.xpath("//button[text()='Edit Contacts']"));
				waitUntilElementClickable("edit", edit, browser, timeout);
				WebElement iconbutton = browser.findElement(By.xpath("//*[contains(@id,'AT1:_ATp:ctb2::icon')]"));
				waitUntilElementClickable("iconbutton", iconbutton, browser, timeout);
				//Thread.sleep(8000);
				
				
				WebElement first = browser.findElement(By.xpath("//label[text()='First Name']//parent::td//following-sibling::td//input[contains(@id,'AT')]"));
				//WebElement first = browser.findElement(By.xpath("(//*[contains(@id,'inputText3::content')])[2]"));
				waitUntilElementClickable("firstname", first, browser, timeout);
				WaituntilElementwritable("firstname", first, browser, First_Name);
				WebElement lastname = browser.findElement(By.xpath("(//*[contains(@id,'inputText4::content')])[2]"));
				waitUntilElementClickable("lastname", lastname, browser, timeout);
				WaituntilElementwritable("lastname", lastname, browser, Last_Name);
				WebElement button1 = browser.findElement(By.xpath("//*[contains(@id,'commandButton1')]"));
				waitUntilElementClickable("button1", button1, browser, timeout);
				WebElement button2 = browser.findElement(By.xpath("(//*[contains(@id,'_ATp:create::icon')])[1]"));
				waitUntilElementClickable("button2", button2, browser, timeout);
				
				
				Select contact = new Select(browser.findElement(By.xpath("//select[contains(@id,'ContactPointType')]")));
				contact.selectByVisibleText(Contact_Point_Type);
				
				//Select contact = new Select(browser.findElement(By.xpath("//*[contains(@id,'ContactPointType::content')]")));
				//Thread.sleep(4000);
				WebElement eamiltext = browser.findElement(By.xpath("//*[contains(@id,'inputText5::content')]"));
				waitUntilElementClickable("eamiltext", eamiltext, browser, timeout);
				WaituntilElementwritable("eamiltext", eamiltext, browser, Email);
				WebElement okbutton1 = browser.findElement(By.xpath("//*[contains(@id,'AT1:dialogOKbtn')]"));
				waitUntilElementClickable("okbutton1", okbutton1, browser, timeout);
				WebElement savebutton = browser.findElement(By.xpath("//button[text()='Save']"));
				waitUntilElementClickable("savebutton", savebutton, browser, timeout);
				By image = By.xpath("//img[contains(@src,'checkmark')]");
				WaituntilElementIsDisplayed("image", image, browser);
				WebElement closebutton = browser.findElement(By.xpath("//button[text()='ave and Close']"));
				waitUntilElementClickable("closebutton", closebutton, browser, timeout);
				//Thread.sleep(6000);
				
				By byTax= By.xpath("(//a[text()='Tax Profile'])[1]");
				browser.findElement(byTax).click();
				
				By byCommunication= By.xpath("(//a[text()='Communication'])[1]");
				browser.findElement(byCommunication).click();
						
				//By email = By.xpath("//span[text()='E-mail']");
				//WaituntilElementIsDisplayed("email", email, browser);				
	
				WaituntilElementIsDisplayed("image", image, browser);
				WebElement saveclosebutton = browser.findElement(By.xpath("//button[text()='ave and Close']"));
				waitUntilElementClickable("saveclosebutton", saveclosebutton, browser, timeout);
				//WebElement done = browser.findElement(By.xpath("//*[contains(@id,'MAnt2:2:AP1:cb1')]"));
				WebElement done = browser.findElement(By.xpath("//button[text()='ne']"));
				
				waitUntilElementClickable("done", done, browser, timeout);
				WebElement homeicon = browser.findElement(By.id("pt1:_UIShome::icon"));
				waitUntilElementClickable("homeicon", homeicon, browser, timeout);
				WebElement receivevalue = browser.findElement(By.linkText("Receivables"));
				waitUntilElementClickable("receivevalue", receivevalue, browser, timeout);
				WebElement billingvalue = browser.findElement(By.linkText("Billing"));
				waitUntilElementClickable("billingvalue", billingvalue, browser, timeout);
				sheet.getRow(i).createCell(15).setCellValue("Pass");
				Updatefile(f, wb);
			}

		} else {
			System.out.println("File is already Processed");
		}
		}
		catch(Exception e)
		{
			e.printStackTrace();
			sheet.getRow(i).createCell(15).setCellValue("Fail");
				Updatefile(f, wb);
				Assert.assertEquals(false, true);
		}
	}

	public static void waitUntilElementClickable(String locatorName, final WebElement elementToWaitFor,
			WebDriver browser, int timeout) {
		System.out.println("<<<<<< " + locatorName + ">>>>>>>>");
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
		System.out.println("<<<<<< " + locatorName + " >>>>>>>>");

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

	public static void waitUntilElementDisappear(String locatorName, final By elementToWaitFor, WebDriver browser) {
		System.out.println("<<<<<< " + locatorName + " >>>>>>>>");
		wait = new WebDriverWait(browser, timeout);
		wait.until(new Function<WebDriver, Boolean>() {
			int j;

			public Boolean apply(WebDriver browser) {
				j++;

				try {
					Boolean flag = browser.findElement(elementToWaitFor).isDisplayed();
					if (flag) {

						return (false);
					} else {
						return (true);
					}

				} catch (Exception e) {
					return false;

				}
			}
		});

	}

	public static void waitUntilElementDisappearCount(String locatorName, final By elementToWaitFor, WebDriver browser) {
		System.out.println("<<<<<< "+locatorName+" >>>>>>>>");
		
		browser.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		wait = new WebDriverWait(browser, timeout);
		wait.until(new Function<WebDriver, Boolean>() {
			int j;
			int k;
		
			public Boolean apply(WebDriver browser) {
				j++;

				try {
					
					List<WebElement> count = browser.findElements(elementToWaitFor);					
					k = count.size();				
				
					if (k==0) {
						return true;
					} else {
						return false;
					}

				} catch (Exception e) {
					return false;

				}
			}
		});
		
		browser.manage().timeouts().implicitlyWait(100, TimeUnit.SECONDS);

	}

	public static void WaituntilElementIsDisplayed(String locatorName, final By elementToWaitFor, WebDriver browser) {
		System.out.println("<<<<<< " + locatorName + " >>>>>>>>");
		wait = new WebDriverWait(browser, timeout);
		wait.until(new Function<WebDriver, Boolean>() {
			int j;

			public Boolean apply(WebDriver browser) {
				j++;

				try {
					Boolean flag = browser.findElement(elementToWaitFor).isDisplayed();
					if (flag) {
						browser.findElement(elementToWaitFor).click();
					} else {
						return (false);
					}

				} catch (Exception e) {
					return false;

				}
				return true;

			}
		});

	}

	public void Updatefile(File f, XSSFWorkbook wb) {
		try {
			FileOutputStream fos = new FileOutputStream(f);
			wb.write(fos);
			fos.flush();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@AfterTest()
	public void Close_Browser() {
//		browser.quit();
	}

}
