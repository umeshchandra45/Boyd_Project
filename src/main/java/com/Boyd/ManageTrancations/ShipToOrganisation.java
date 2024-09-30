package com.Boyd.ManageTrancations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class ShipToOrganisation extends Base_Class{
	//public static WebDriverWait wait;
	//public static int timeout = 60;
	public String item;
	public String supplier;
	@Test
	public void homepage() throws Exception{
		File f=new File(System.getProperty("user.dir")+"\\Excel\\ShipToOrg.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Sheet1");
		int totalrows = sheet.getPhysicalNumberOfRows();
		WebElement tasks = browser.findElement(By.xpath("//img[contains(@title,'Tasks')]"));
		waitUntilElementClickable("tasks", tasks, browser, timeout);
		Thread.sleep(13000);
		WebElement ManageApprovedSupplierList = browser.findElement(By.xpath("//a[text()='Manage Approved Supplier List Entries']"));
		waitUntilElementClickable("ManageApprovedSupplierList", ManageApprovedSupplierList, browser, timeout);
		WebElement add_img = browser.findElement(By.xpath("//img[contains(@id,'AT2:_ATp:create::icon')]"));
		waitUntilElementClickable("add_img", add_img, browser, timeout);
		Thread.sleep(8000);
		for(int i=1;i<totalrows;i++)
		{
			System.out.println("Count of i value :" +i);
			DataFormatter df = new DataFormatter();
			item = df.formatCellValue(sheet.getRow(i).getCell(1));
			System.out.println(item);
			supplier=sheet.getRow(i).getCell(2).getStringCellValue();
			System.out.println(supplier);
			WebElement ele2 = browser.findElement(By.xpath("//select[contains(@id,'AP1:prcBuId::content')]"));
			Select typ1 = new Select(ele2);
			typ1.selectByVisibleText("IN VAD");
			Thread.sleep(3000);
			WebElement ele3 = browser.findElement(By.xpath("//select[contains(@id,'AP1:ScopeId::content')]"));
			Select typ = new Select(ele3);
			typ.selectByVisibleText("Ship-to Organization");
			Thread.sleep(1000);
			WebElement btn_yes = browser.findElement(By.xpath("//button[text()='Yes']"));
			waitUntilElementClickable("btn_yes", btn_yes, browser, timeout);
			Thread.sleep(5000);
			//		WebElement in_ShipOrg = browser.findElement(By.xpath("//input[contains(@id,'AP1:ShipToOrganization::content')]"));
			//		waitUntilElementClickable("in_ShipOrg", in_ShipOrg, browser, timeout);
			//		in_ShipOrg.clear();
			//		in_ShipOrg.sendKeys("Vadodara");
			WebElement a_item = browser.findElement(By.xpath("//a[contains(@id,'AP1:Item::lovIconId')]"));
			waitUntilElementClickable("a_item", a_item, browser, timeout);
			WebElement in_item = browser.findElement(By.xpath("//input[contains(@id,'AP1:Item::_afrLovInternalQueryId:value00::content')]"));
			waitUntilElementClickable("in_item", in_item, browser, timeout);
			in_item.clear();
			in_item.sendKeys(item);
			WebElement search_item = browser.findElement(By.xpath("//button[text()='Search']"));
			waitUntilElementClickable("search_item", search_item, browser, timeout);
			try{
				WebElement select_item = browser.findElement(By.xpath("//*[contains(@id,'AP1:Item_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
				select_item.click();
				System.out.println("item selected sucessfully");
				WebElement ok_item = browser.findElement(By.xpath("//button[contains(@id,'AP1:Item::lovDialogId::ok')]"));
				waitUntilElementClickable("ok_item", ok_item, browser, timeout);
//				WebElement a_supplier = browser.findElement(By.xpath("//a[contains(@id,'AP1:Supplier::lovIconId')]"));
//				waitUntilElementClickable("a_supplier", a_supplier, browser, timeout);
//				WebElement in_supplier = browser.findElement(By.xpath("//input[contains(@id,'Supplier::_afrLovInternalQueryId:value00::content')]"));
//				waitUntilElementClickable("in_supplier", in_supplier, browser, timeout);
//				in_supplier.clear();
//				in_supplier.sendKeys(supplier);
//				in_supplier.sendKeys(Keys.ENTER);
//				Thread.sleep(3000);
//				WebElement btw_SearchSupplier = browser.findElement(By.xpath("//button[contains(@id,'Supplier::_afrLovInternalQueryId::search')]"));
//				waitUntilElementClickable("btw_SearchSupplier", btw_SearchSupplier, browser, timeout);
//				WebElement td_supplier = browser.findElement(By.xpath("//div[contains(@id,'AP1:Supplier_afrLovInternalTableId::db')]/table/tbody/tr/td[1]"));
//				waitUntilElementClickable("td_supplier", td_supplier, browser, timeout);
//				WebElement Ok_supplier = browser.findElement(By.xpath("//button[contains(@id,'AP1:Supplier::lovDialogId::ok')]"));
//				waitUntilElementClickable("Ok_supplier", Ok_supplier, browser, timeout);
							WebElement in_supplier = browser.findElement(By.xpath("//input[contains(@id,':AP1:Supplier::content')]"));
							waitUntilElementClickable("in_supplier", in_supplier, browser, timeout);
							in_supplier.clear();
							in_supplier.sendKeys(supplier);
							in_supplier.sendKeys(Keys.ENTER);
				Thread.sleep(3000);
				WebElement a_SaveNclose = browser.findElement(By.xpath("//a[@title='Save and Close']"));
				waitUntilElementClickable("a_SaveNclose", a_SaveNclose, browser, timeout);
				WebElement td_SaveNcreateanother = browser.findElement(By.xpath("//td[text()='Save and Create Another']"));
				waitUntilElementClickable("td_SaveNcreateanother", td_SaveNcreateanother, browser, timeout);
				Thread.sleep(1000);
				WebElement btn_ok = browser.findElement(By.xpath("//button[@id='d1::msgDlg::cancel']"));
				waitUntilElementClickable("btn_ok", btn_ok, browser, timeout);
				System.out.println("Created sucessfully");
				System.out.println("====================================================");
				sheet.getRow(i).createCell(3).setCellValue("pass");
				Updatefile(f, wb);
				Thread.sleep(3000);
			} 
			catch(Exception e)
			{
				if (e.getMessage().contains("Item")) {
					WebElement ok_item = browser.findElement(By.xpath("//button[contains(@id,'AP1:Item::lovDialogId::ok')]"));
					waitUntilElementClickable("ok_item", ok_item, browser, timeout);
					WebElement btn_cancel = browser.findElement(By.xpath("//span[text()='ancel']"));
					waitUntilElementClickable("btn_cancel", btn_cancel, browser, timeout);
					Thread.sleep(3000);
					WebElement add_img2 = browser.findElement(By.xpath("//img[contains(@id,'AT2:_ATp:create::icon')]"));
					waitUntilElementClickable("add_img2", add_img2, browser, timeout);
					System.out.println("item not found");
					System.out.println("====================================================");
					sheet.getRow(i).createCell(3).setCellValue("Failed");
					sheet.getRow(i).createCell(4).setCellValue("item not found");
					Updatefile(f, wb);
					Thread.sleep(4000);
				} else if (e.getMessage().contains("Supplier")) {
					WebElement Ok_supplier = browser.findElement(By.xpath("//button[contains(@id,'AP1:Supplier::lovDialogId::ok')]"));
					waitUntilElementClickable("Ok_supplier", Ok_supplier, browser, timeout);
					WebElement btn_cancel = browser.findElement(By.xpath("//span[text()='ancel']"));
					waitUntilElementClickable("btn_cancel", btn_cancel, browser, timeout);
					Thread.sleep(3000);
					WebElement add_img2 = browser.findElement(By.xpath("//img[contains(@id,'AT2:_ATp:create::icon')]"));
					waitUntilElementClickable("add_img2", add_img2, browser, timeout);
					System.out.println("Supplier not found");
					System.out.println("====================================================");
					sheet.getRow(i).createCell(3).setCellValue("Failed");
					sheet.getRow(i).createCell(4).setCellValue("supplier not found");
					Updatefile(f, wb);
					Thread.sleep(4000);
				}
				else {
					System.out.println("An unexpected error occurred: " + e.getMessage());

				}
			}
		}
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

