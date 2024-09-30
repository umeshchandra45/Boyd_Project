package com.Boyd.ManageTrancations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
public class UserAddOn2 extends Base_Class {
	public static WebDriverWait wait;
	public static int timeout = 60;
	public String role;
	public String userName;
	public WebElement equaltoUserName(String eqalto2) {
        String xpath = "//td[text()='" + eqalto2 + "']";
        return browser.findElement(By.xpath(xpath));
    }
	@Test
	public void homepage() throws Exception {
		File f=new File(System.getProperty("user.dir")+"\\Excel\\UserAddOn.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Sheet3");
		int totalrows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of Excel rows are :" +totalrows);
		String a = "";
		wait=new WebDriverWait(browser, 50 , 500);
		for(int i=1;i<totalrows;i++)
		{
		System.out.println("Count of i value :" +i);
		role=sheet.getRow(i).getCell(0).getStringCellValue();
		if(a == role ){
			userName=sheet.getRow(i).getCell(1).getStringCellValue();
			WebElement inputValue =  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:usSrcBx::content')]")));
			waitUntilElementClickable("inputValue", inputValue, browser, timeout);
			inputValue.clear();
			inputValue.sendKeys(userName); 
			Thread.sleep(1000);
			WebElement search1 =  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'uSp1:cil1::icon')]")));
			waitUntilElementClickable("search1", search1, browser, timeout);
			try {
				equaltoUserName(userName).isDisplayed();
	
				/*equaltoUserName(userName).click();
			    Thread.sleep(6000);
				WebElement addUserToRole = browser.findElement(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:usAddUs')]"));
				if(addUserToRole.isEnabled()) {
				waitUntilElementClickable("addUserToRole", addUserToRole, browser, timeout);
				Thread.sleep(2000);*/
				sheet.getRow(i).createCell(2).setCellValue("displayed");
				Updatefile(f, wb);
			/*
				}
				else {
				sheet.getRow(i).createCell(2).setCellValue("not displayed");
				Updatefile(f, wb);
				System.out.println("user not found");}*/
			}
			catch(Exception e)
			{
				sheet.getRow(i).createCell(2).setCellValue("not found");
				Updatefile(f, wb);
				System.out.println("user not found");
			}
					
			}
		else {
        if(a!="")
        {
        	WebElement close = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:d1::close')]")));
    		waitUntilElementClickable("close", close, browser, timeout);
    		WebElement next = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:cb4')]")));
    		waitUntilElementClickable("next", next, browser, timeout);
    		WebElement saveAndClose = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:3:sSp1:cb1')]")));
    		waitUntilElementClickable("saveAndClose", saveAndClose, browser, timeout);
    		//*[contains(@id,'_FOd1::msgDlg::cancel')]
    		WebElement ok = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOd1::msgDlg::cancel')]")));
    		waitUntilElementClickable("ok", ok, browser, timeout);
        }
        a = role;
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
		WebElement addUser = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:cil3')]")));
		waitUntilElementClickable("addUser", addUser, browser, timeout);		
		userName=sheet.getRow(i).getCell(1).getStringCellValue();
		WebElement inputValue = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:usSrcBx::content')]")));
		waitUntilElementClickable("inputValue", inputValue, browser, timeout);
		inputValue.clear();
		inputValue.sendKeys(userName); 
		Thread.sleep(1000); 
		WebElement search1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'uSp1:cil1::icon')]")));
		waitUntilElementClickable("search1", search1, browser, timeout);
		try {
			equaltoUserName(userName).isDisplayed();

			/*equaltoUserName(userName).click();
		    Thread.sleep(6000);
			WebElement addUserToRole = browser.findElement(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:usAddUs')]"));
			if(addUserToRole.isEnabled()) {
			waitUntilElementClickable("addUserToRole", addUserToRole, browser, timeout);
			Thread.sleep(2000);*/
			sheet.getRow(i).createCell(2).setCellValue("displayed");
			Updatefile(f, wb);
		/*
			}
			else {
			sheet.getRow(i).createCell(2).setCellValue("not displayed");
			Updatefile(f, wb);
			System.out.println("user not found");}*/
		}
		catch(Exception e)
		{
			sheet.getRow(i).createCell(2).setCellValue("not found");
			Updatefile(f, wb);
			System.out.println("user not found");
		}
		}
}
		Thread.sleep(2000);
    	WebElement close = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:d1::close')]")));
		waitUntilElementClickable("close", close, browser, timeout);
		WebElement next = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:2:uSp1:cb4')]")));
		waitUntilElementClickable("next", next, browser, timeout);
		WebElement saveAndClose = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOpt1:_FOr1:0:_FOSrASE_FUSE_SECURITY_CONSOLE:0:MAnt2:3:sSp1:cb1')]")));
		waitUntilElementClickable("saveAndClose", saveAndClose, browser, timeout);
		WebElement ok = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@id,'_FOd1::msgDlg::cancel')]")));
		waitUntilElementClickable("ok", ok, browser, timeout);
		System.out.println("excuted");
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
