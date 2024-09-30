package com.Boyd.ManageTrancations;

import java.util.concurrent.TimeUnit;
import java.util.function.Function;

import org.openqa.selenium.By;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Base_Class {
	public WebDriver browser;
	public static WebDriverWait wait;
	public static int timeout = 60;

	@BeforeMethod
	public void Login_Page() throws InterruptedException {

		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(20, TimeUnit.SECONDS);
		browser.get("https://elme.fa.us8.oraclecloud.com/");
	//	Thread.sleep(2000);
	//	browser.get("https://elme-dev2.fa.us8.oraclecloud.com");
		//	browser.get("https://elme-test.login.us8.oraclecloud.com");
		Thread.sleep(7000);
		browser.findElement(By.id("userid")).click();
		Thread.sleep(1000);
		browser.findElement(By.id("userid")).sendKeys("forsys.user");
		browser.findElement(By.id("password")).click();
		Thread.sleep(1000);
		browser.findElement(By.id("password")).sendKeys("forsys4@4!");
		Thread.sleep(1000);
	//	browser.findElement(By.id("password")).sendKeys("welcome1");
		browser.findElement(By.id("btnActive")).click();
		Thread.sleep(19000);
		wait=new WebDriverWait(browser, 40 , 3000);
		WebElement homepage = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[text()='You have a new home page!']")));
		waitUntilElementClickable("homepage", homepage, browser, timeout);
		//email updation
		Thread.sleep(19000);
		WebElement receivables = browser.findElement(By.linkText("Receivables"));
		waitUntilElementClickable("receivables", receivables, browser, timeout);
		WebElement AcccountReceivable = browser.findElement(By.linkText("Accounts Receivable"));
		waitUntilElementClickable("AcccountReceivable", AcccountReceivable, browser, timeout);
		//ship to organisation
		/*Thread.sleep(5000);
		WebElement procurement = browser.findElement(By.linkText("Procurement"));
		waitUntilElementClickable("procurement", procurement, browser, timeout);
		WebElement po = browser.findElement(By.linkText("Purchase Orders"));
		waitUntilElementClickable("po", po, browser, timeout);
		Thread.sleep(5000);*/
		/*	//(for intercompany) and (customer creation)
		Thread.sleep(10000);
		WebElement receivables = browser.findElement(By.linkText("Receivables"));
		waitUntilElementClickable("receivables", receivables, browser, timeout);
		WebElement billing = browser.findElement(By.linkText("Billing"));
		waitUntilElementClickable("billing", billing, browser, timeout);*/
		//(for user addition)
		/*	WebElement tools = wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Tools")));
		waitUntilElementClickable("tools", tools, browser, timeout);
		WebElement security = wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Security Console")));
		waitUntilElementClickable("security", security, browser, timeout);*/
		//customer creation
	

	}

	@AfterMethod
	public void Quit_browser()
	{
		//		browser.quit();
	}

	public static void waitUntilElementClickable(String locatorName, final WebElement elementToWaitFor,
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

	public static void WaituntilElementwritable(String locatorName, final WebElement elementToWaitFor,
			WebDriver browser, String value) {
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

}
