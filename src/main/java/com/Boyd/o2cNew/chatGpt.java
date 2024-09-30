package com.Boyd.o2cNew;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class chatGpt {
	
 public static void main(String[] args) {
	 WebDriverManager.chromedriver().setup();
	 WebDriver driver = new ChromeDriver();
	 driver.get("https://platform.openai.com/");
	 WebElement loginButton= driver.findElement(By.xpath("//div[text()='Log in']"));
	 loginButton.click();
	 WebElement usernameField = driver.findElement(By.name("email"));
	 usernameField.sendKeys("umeshchandrareddy123@gmail.com");
	 driver.findElement(By.xpath("//button[text()='Continue']")).click();
	 WebElement passwordField = driver.findElement(By.name("password"));
	 passwordField.sendKeys("Skype!chatgpt45");
     driver.findElement(By.xpath("//button[text()='Continue']")).click();
	 
}
}
