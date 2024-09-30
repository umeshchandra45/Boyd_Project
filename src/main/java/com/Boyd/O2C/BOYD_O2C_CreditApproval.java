package com.Boyd.O2C;

import java.util.concurrent.TimeUnit;
import java.util.function.Function;

import org.openqa.selenium.By;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BOYD_O2C_CreditApproval {
	public WebDriver browser;
	public static int timeout = 60;
	public static WebDriverWait wait;
	public String ordernumber="66123";
	@BeforeTest
	public void login_test() throws InterruptedException {
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(80, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(80, TimeUnit.SECONDS);
		browser.get("https://elme-dev1.fa.us8.oraclecloud.com");
		browser.findElement(By.id("userid")).click();
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		browser.findElement(By.id("password")).sendKeys("forsys2023");
		browser.findElement(By.id("btnActive")).click();
		WebDriverWait wait1 = new WebDriverWait(browser, 500, 20);
		wait1.until(ExpectedConditions.elementToBeClickable(By.id("pt1:_UIShome")));
		browser.findElement(By.id("pt1:_UIShome")).click();
		WebElement receivables = browser.findElement(By.xpath("//a[contains(text(),'Receivables')]"));
		waitUntilElementClickable(receivables, browser, timeout);
		WebElement creditReview = browser.findElement(By.xpath("//a[contains(text(),'Credit Reviews')]"));
		waitUntilElementClickable(creditReview, browser, timeout);

	}
	@Test
	public void creditApprove() throws InterruptedException {
		WebElement showFilter = browser.findElement(By.xpath("//a[contains(text(),'Show Filters')]"));
		waitUntilElementClickable(showFilter, browser, timeout);
		Thread.sleep(6000);
		WebElement savedSearch = browser.findElement(By.xpath("//select[contains(@id, 'LSQrySS::content')]"));
		Select filter = new Select(savedSearch);
		filter.selectByVisibleText("Case Folders Pending My Approval");
		Thread.sleep(3000);
		WebElement status = browser.findElement(By.xpath("//a[contains(@id, 'value30::drop')]"));
		waitUntilElementClickable(status, browser, timeout);
		Thread.sleep(3000);
		WebElement status2 = browser.findElement(By.xpath("//label[contains(text(), 'New')]"));
		waitUntilElementClickable(status2, browser, timeout);
		Thread.sleep(2000);
		WebElement approvar = browser.findElement(By.xpath("//input[contains(@id, 'LSQry:value70::content')]"));
		waitUntilElementClickable(approvar, browser, timeout);
		approvar.clear();
		Thread.sleep(2000);
		WebElement search = browser.findElement(By.xpath("//button[contains(text(), 'Search')]"));
		waitUntilElementClickable(search, browser, timeout);
		Thread.sleep(2000);
		WebElement table = browser.findElement(By.xpath("(//div[contains(@id,'_ATp:ATt1::db')]//table//td)[1]"));
		waitUntilElementClickable(table, browser, timeout);
		Thread.sleep(2000);
		WebElement Reassign = browser.findElement(By.xpath("//button[contains(text(),'Reassign')]"));
		waitUntilElementClickable(Reassign, browser, timeout);
		WebElement creditAnalyst = browser.findElement(By.xpath("//a[contains(@id,'creditAnalystTempId::lovIconId')]"));
		waitUntilElementClickable(creditAnalyst, browser, timeout);
		WebElement searchicon = browser.findElement(By.linkText("Search..."));
		waitUntilElementClickable( searchicon, browser, timeout);
		WebElement name = browser.findElement(By.xpath("//*[contains(@id,'afrLovInternalQueryId:value00::content')]"));
		waitUntilElementClickable(name, browser, timeout);
		name.sendKeys("Forsys");
		WebElement searchbutton = browser.findElement(By.xpath("//*[contains(@id,'_afrLovInternalQueryId::search')]"));
		waitUntilElementClickable(searchbutton, browser, timeout);
		WebElement table1 = browser.findElement(By.xpath("//*[contains(@id,'_afrLovInternalTableId::db')]/table/tbody/tr[1]/td[1]"));
		waitUntilElementClickable(table1, browser, timeout);
		WebElement okbutton = browser.findElement(By.xpath("//*[contains(@id,'lovDialogId::ok')]"));
		waitUntilElementClickable(okbutton, browser, timeout);
		WebElement saveAndClose = browser.findElement(By.xpath("//button[contains(text(),'ave and Close')]"));
		waitUntilElementClickable(saveAndClose, browser, timeout);
		Thread.sleep(5000);
		WebElement number = browser.findElement(By.xpath("//a[contains(@id,'FOTsr1:0:SP2:ls1:AT1:_ATp:ATt1:0:cl1')]"));
		waitUntilElementClickable(number, browser, timeout);
		WebElement Recommendations = browser.findElement(By.xpath("(//a[contains(text(),'Recommendations')])[1]"));
		waitUntilElementClickable(Recommendations, browser, timeout);
		WebElement type = browser.findElement(By.xpath("//select[contains(@id, 'ATp:tab1:0:socRecommCode::content')]"));
		Select type1 = new Select(type);
		type1.selectByVisibleText("Approve Source Transaction Credit Request");
		Thread.sleep(3000);
		WebElement save = browser.findElement(By.xpath("//span[text()='Save']"));
		waitUntilElementClickable(save, browser, timeout);
		Thread.sleep(3000);
		WebElement Actions = browser.findElement(By.xpath("//a[contains(text(),'Actions')]"));
		waitUntilElementClickable(Actions, browser, timeout);
		Thread.sleep(2000);
		WebElement approve = browser.findElement(By.xpath("//td[contains(text(),'Approve')]"));
		waitUntilElementClickable(approve, browser, timeout);
		WebElement textbox = browser.findElement(By.xpath("//*[contains(@id,'credit:0:_FOTsr1:1:AP1:it3::content')]"));
		waitUntilElementClickable(textbox, browser, timeout);
		textbox.sendKeys("Approved");
		Thread.sleep(3000);
		WebElement saveAndClose1 = browser.findElement(By.xpath("//button[contains(text(),'ave and Close')]"));
		waitUntilElementClickable(saveAndClose1, browser, timeout);
		Thread.sleep(3000);
		WebElement done = browser.findElement(By.xpath("//span[text()='one']"));
		waitUntilElementClickable(done, browser, timeout);
		Thread.sleep(2000);
		WebElement home = browser.findElement(By.id("pt1:_UIShome::icon"));
		waitUntilElementClickable(home, browser, timeout);
		Thread.sleep(2000);
		WebElement ordermanagement = browser.findElement(By.linkText("Order Management"));
		waitUntilElementClickable( ordermanagement, browser, timeout);
		WebElement order1 = browser.findElement(By.id("itemNode_order_management_order_management_1"));
		waitUntilElementClickable(order1, browser, timeout);
		Thread.sleep(10000);
		WebElement input_order = browser.findElement(By.xpath("//input[contains(@id,'FOTsr1:0:AP1:qq1:it1::content')]"));
		waitUntilElementClickable(input_order, browser, timeout);
		input_order.sendKeys(ordernumber);
		Thread.sleep(10000);
		WebElement seach_icon = browser.findElement(By.xpath("//a[contains(@id,'FOTsr1:0:AP1:qq1::search_icon')]"));
		waitUntilElementClickable(seach_icon, browser, timeout);
		Thread.sleep(7000);
		WebElement Refresh = browser.findElement(By.xpath("//button[contains(text(),'Refresh')]"));
		waitUntilElementClickable(Refresh, browser, timeout);
		Thread.sleep(6000);
		WebElement Refresh1 = browser.findElement(By.xpath("//button[contains(text(),'Refresh')]"));
		waitUntilElementClickable(Refresh1, browser, timeout);
		Thread.sleep(6000);
		WebElement Refresh2 = browser.findElement(By.xpath("//button[contains(text(),'Refresh')]"));
		waitUntilElementClickable(Refresh2, browser, timeout);
		Thread.sleep(6000);
		WebElement Refresh3 = browser.findElement(By.xpath("//button[contains(text(),'Refresh')]"));
		waitUntilElementClickable(Refresh3, browser, timeout);
		Thread.sleep(6000);
		WebElement Refresh4 = browser.findElement(By.xpath("//button[contains(text(),'Refresh')]"));
		waitUntilElementClickable(Refresh4, browser, timeout);
		Thread.sleep(6000);
		WebElement Refresh5 = browser.findElement(By.xpath("//button[contains(text(),'Refresh')]"));
		waitUntilElementClickable(Refresh5, browser, timeout);

		
		
		
	}
	public static void waitUntilElementClickable(final WebElement elementToWaitFor,
			WebDriver browser, int timeoutsec) {
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
}
