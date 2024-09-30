package com.Boyd.ManageTrancations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class UserConfirmation extends Base_Class{
	public static WebDriverWait wait;
	public static int timeout = 60;
	public String role;
	public String userName;
	@Test
	public void homepage() throws Exception {
		File f=new File(System.getProperty("user.dir")+"\\Excel\\UserConfirmation.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Sheet1");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		wait=new WebDriverWait(browser, 50 , 500);
			role="BYD DC and Procurement Employee Custom Req";
			WebElement roles = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:_FOTr0:0:sp1:srchBox::content')]")));
			waitUntilElementClickable("roles", roles, browser, timeout);
			roles.clear();
			WaituntilElementwritable("roles", roles, browser, role);
			WebElement search = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:_FOTr0:0:sp1:cil1::icon')]")));
			waitUntilElementClickable("search", search, browser, timeout);
			WebElement actions = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:_FOTr0:0:sp1:resList:0:cb1')]")));
			waitUntilElementClickable("actions", actions, browser, timeout);
			WebElement editRole = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:_FOTr0:0:sp1:resList:0:cmiEdit')]/td[2]")));
			waitUntilElementClickable("editRole", editRole, browser, timeout);
			WebElement user = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[contains(@title,'Users Step: Not Visited Step')]")));
			waitUntilElementClickable("user", user, browser, timeout);
			for(int i=1;i<totalrows;i++)
			{
				System.out.println("Count of i value :" +i);
				userName=sheet.getRow(i).getCell(0).getStringCellValue();
			WebElement input = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[contains(@id,'rmUsrT:Ase1Ip::content')]")));
			waitUntilElementClickable("input", input, browser, timeout);
			input.clear();
			input.click();
			input.sendKeys(userName);
			input.sendKeys(Keys.ENTER);
			Thread.sleep(3000);
			try {
				String xpath1 = "//span[text()='" + userName + "']";
				browser.findElement(By.xpath(xpath1));
				System.out.println(browser.findElement(By.xpath(xpath1)).getText());
				sheet.getRow(i).createCell(1).setCellValue("displayed");
				Updatefile(f, wb);
				
			}
			catch(Exception e)
			{
				System.out.println("usernotfound");
				sheet.getRow(i).createCell(1).setCellValue("not displayed");
				Updatefile(f, wb);
			}		
			}
			System.out.println("file execution completed");
			try
			{
				wb.close();
			}
			catch(Exception e)
			{
				
			}
			
			
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
}