package Framework;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import io.github.bonigarcia.wdm.WebDriverManager;

@SuppressWarnings("unused")
public class Base_Test {
	
	public ExtentHtmlReporter htmlReporter;
	public ExtentReports extent;
	public ExtentTest extentTest;
	public static WebDriver browser;
	
	
	@BeforeMethod()
	public void Start_Up()
	{
		String timeStamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new Date());
		htmlReporter = new ExtentHtmlReporter(System.getProperty("user.dir")+"\\Extent\\"+timeStamp+" test_results.html");
		htmlReporter.config().setEncoding("utf-8");
		htmlReporter.config().setDocumentTitle("Automation Reports");
		htmlReporter.config().setReportName("Automation Test Results");
		htmlReporter.config().setTheme(Theme.STANDARD);

		extent = new ExtentReports();
		extent.setSystemInfo("Browser", "Chrome");
		extent.attachReporter(htmlReporter);
	
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.setPageLoadStrategy(PageLoadStrategy.NONE);
		browser = new ChromeDriver(options);
		browser.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		browser.manage().window().maximize();
	}
	
	public String captureScreen() throws IOException {
		TakesScreenshot screen = (TakesScreenshot) browser;
		File src = screen.getScreenshotAs(OutputType.FILE);
		String dest ="D:\\Srikanth Workspace\\Harmonic_Boyd\\Screenshot\\ScreenShots"+System.currentTimeMillis()+".png";
		File target = new File(dest);
		FileUtils.copyFile(src, target);
		return dest;
	}

	
	@AfterMethod()
	public void Close_Browser()
	{
//		browser.quit();
		extent.flush();
	}
}
