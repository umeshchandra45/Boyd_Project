package com.Boyd.O2C.prac;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class scroll {

	public WebDriver driver;
	public static WebDriverWait wait;
	public static int timeout = 60;
	public void Login_Page() throws InterruptedException {

		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		driver = new ChromeDriver(options);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(20, TimeUnit.SECONDS);
		driver.get("Quotation: Q-00908442 ~ Salesforce - Unlimited Edition");
		
}
	public static void main(String[] args) throws Exception {
		scroll scroll1=new scroll();
		scroll1.Login_Page();
	}}
