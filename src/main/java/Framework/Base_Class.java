package Framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

import io.github.bonigarcia.wdm.WebDriverManager;

@SuppressWarnings("unused")
public class Base_Class {
	public ExtentHtmlReporter htmlReporter;
	public ExtentReports extent;
	public ExtentTest extentTest;
   public static WebDriver browser;
	
	@BeforeMethod()
	public  void Login_Page() throws Exception
	{
		
//		htmlReporter = new ExtentHtmlReporter(System.getProperty("user.dir")+"\\Extent\\test_result.html");
//		htmlReporter.config().setEncoding("utf-8");
//		htmlReporter.config().setDocumentTitle("Automation Reports");
//		htmlReporter.config().setReportName("Automation Test Results");
//		htmlReporter.config().setTheme(Theme.STANDARD);
//		
//		extent = new ExtentReports();
//		extent.setSystemInfo("Browser", "Chrome");
//		extent.attachReporter(htmlReporter);
		
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().window().maximize();
		browser.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		browser.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
		File f = new File(System.getProperty("user.dir")+"\\Propertyfile\\Config.properties");
		FileInputStream fis = new FileInputStream(f);
		Properties prop = new Properties();
		prop.load(fis);
		browser.get(prop.getProperty("url"));
		WebElement name = browser.findElement(By.id("userid"));
		highLightElement(browser, name);
		name.sendKeys(prop.getProperty("username"));
		WebElement pasword = browser.findElement(By.id("password"));
		highLightElement(browser, pasword);
		pasword.sendKeys(prop.getProperty("password"));
		WebElement signin = browser.findElement(By.id("btnActive"));
		highLightElement(browser, signin);
		signin.click();
		Thread.sleep(5000);
		WebElement button = browser.findElement(By.linkText("You have a new home page!"));
		highLightElement(browser, button);
		button.click();
		Thread.sleep(3000);
		WebElement receivable = browser.findElement(By.linkText("Receivables"));
		highLightElement(browser, receivable);
		receivable.click();
		WebElement billing = browser.findElement(By.linkText("Billing"));
		highLightElement(browser, billing);
		billing.click();
		Thread.sleep(4000);
		WebElement task = browser.findElement(By.xpath("//img[contains(@title,'Tasks')]"));
		highLightElement(browser, task);
		task.click();
		WebElement transaction = browser.findElement(By.linkText("Create Transaction"));
		highLightElement(browser, transaction);
		transaction.click();
		Thread.sleep(6000);
	}
	
	public static void highLightElement(WebDriver browser,WebElement ele)
	{
		try {  
            JavascriptExecutor js = (JavascriptExecutor) browser;  
            js.executeScript("arguments[0].style.border='4px groove green'", ele);
            Thread.sleep(1000);  
            js.executeScript("arguments[0].style.border=''", ele);  
       } catch (Exception e) {  
            System.out.println(e);  
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
	
	
	
	
	@AfterMethod()
	public void Quit_Browser()
	{
//		browser.quit();
//		extent.flush();
	}

}
