package com.selenium.automation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.util.regex.Pattern;
import java.util.concurrent.TimeUnit;

import org.junit.*;
import java.beans.Introspector;

import static org.junit.Assert.*;
import static org.hamcrest.CoreMatchers.*;

import org.openqa.selenium.*;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Main {
	
	  private static WebDriver driver = new FirefoxDriver();
	  
	  private static int numberofColumns;
	  private static int y;
	  private static String firstName;
	  private static String lastName;
	  private static String userNameEmail;
	  private static String schoolType;
	  private static String year;
	  private static String passwordName;
	  private static String categoryName;
	  private static String pitchTitle;
	  private static String pitchDescription;
	  private static String pathName;
	  private static Sheet finalSheet;
	  
	  
	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		
		FileInputStream input1 = new FileInputStream("/Users/neel/Desktop/uploadDatabase.xls");
		Workbook book1 = Workbook.getWorkbook(input1);
		Sheet Sheet1=book1.getSheet("Sheet1");
		finalSheet = Sheet1;
		
		y = 0;
		getNumberOfColumns();
	 	
	 	while(y < numberofColumns)
	 	{
	 		processPitchNewUser();
	 	}
    }
	
	public static void writeData() throws RowsExceededException, WriteException, IOException, BiffException
	{
		FileInputStream userInput = new FileInputStream("/Users/neel/Desktop/userdatabase.xls");
		Workbook bookUserDataBase = Workbook.getWorkbook(userInput);
		Sheet userSheet = bookUserDataBase.getSheet("Sheet1");
		
		FileOutputStream outputStream = new FileOutputStream("/Users/neel/Desktop/output.xls");
		WritableWorkbook outputWorkbook = Workbook.createWorkbook(outputStream);
		WritableSheet outputSheet = outputWorkbook.createSheet("resultautomation", 0);
		int userDataBaseSheet = userSheet.getRows();
		
		for(int i=0; i < userDataBaseSheet; i++)
		{
			userNameEmail = userSheet.getCell(0, i).getContents();
			passwordName = userSheet.getCell(1, i).getContents();
			
			driver.get("http://1000pitches.com");
		    driver.findElement(By.linkText("Login")).click();
		    driver.findElement(By.id("edit-name")).clear();
		    driver.findElement(By.id("edit-name")).sendKeys(userNameEmail);
		    driver.findElement(By.id("edit-pass")).clear();
		    driver.findElement(By.id("edit-pass")).sendKeys(passwordName);
		    driver.findElement(By.id("edit-submit")).click();
		    try
		    {
		    	driver.findElement(By.xpath("//li[3]/a/span")).click();
		    }
		    catch(NoSuchElementException exception)
		    {
		    	driver.findElement(By.id("edit-pass")).clear();
		    	passwordName = Introspector.decapitalize(passwordName);
			    driver.findElement(By.id("edit-pass")).clear();
			    driver.findElement(By.id("edit-pass")).sendKeys(passwordName);
			    driver.findElement(By.id("edit-submit")).click();
			    try
			    {
			    	driver.findElement(By.xpath("//li[3]/a/span")).click();
			    }
			    catch(NoSuchElementException exception2)
			    {
			    	outputWorkbook.write();
			    	outputWorkbook.close();
			    }
		    }
		    driver.findElement(By.cssSelector("p")).click();
		    String pT = driver.findElement(By.cssSelector("lh")).getText();
		    String pD = driver.findElement(By.cssSelector("dd")).getText();
		    
		    System.out.println(pT);
		    
		    Label data1 = new Label(0, i, pT);
	        outputSheet.addCell(data1);

	        Label data2 = new Label(1, i, pD);
	        outputSheet.addCell(data2);
	        driver.findElement(By.linkText("Logout")).click();
		}
		outputWorkbook.write();
		outputWorkbook.close();
	}
	
	public static void getNumberOfColumns() throws BiffException, IOException
	{	
			numberofColumns = finalSheet.getColumns();
	}
	
	public static void processPitchExistingUser() throws BiffException, IOException
	{
		userNameEmail = finalSheet.getCell(0, y).getContents();
		passwordName = finalSheet.getCell(1, y).getContents();
		categoryName = finalSheet.getCell(2, y).getContents();
		pitchTitle = finalSheet.getCell(3, y).getContents();
		pitchDescription = finalSheet.getCell(4, y).getContents();
		pathName = finalSheet.getCell(5, y).getContents();
				
		System.out.println(userNameEmail);
		System.out.println(passwordName);
		System.out.println(categoryName);
		System.out.println(pitchTitle);
		System.out.println(pitchDescription);
		System.out.println(pathName);
		
		y++;
		
		driver.get("http://1000pitches.com");
	    driver.findElement(By.linkText("Login")).click();
	    driver.findElement(By.id("edit-name")).clear();
	    driver.findElement(By.id("edit-name")).sendKeys(userNameEmail);
	    driver.findElement(By.id("edit-pass")).clear();
	    driver.findElement(By.id("edit-pass")).sendKeys(passwordName);
	    driver.findElement(By.id("edit-submit")).click();
	    try
	    {
	    	driver.findElement(By.xpath("//li[3]/a/span")).click();
	    }
	    catch(NoSuchElementException exception)
	    {
	    	driver.findElement(By.id("edit-pass")).clear();
	    	passwordName = Introspector.decapitalize(passwordName);
		    driver.findElement(By.id("edit-pass")).clear();
		    driver.findElement(By.id("edit-pass")).sendKeys(passwordName);
		    driver.findElement(By.id("edit-submit")).click();
		    driver.findElement(By.xpath("//li[3]/a/span")).click();
	    }
	    
	    System.out.println(userNameEmail);
	    
	    driver.findElement(By.linkText("Pitch Your Idea")).click();
	    driver.findElement(By.id("edit-title")).clear();
	    new Select(driver.findElement(By.id("edit-category"))).selectByVisibleText(categoryName);
	    driver.findElement(By.id("edit-title")).sendKeys(pitchTitle);
	    driver.findElement(By.id("edit-station")).click();
	    driver.findElement(By.id("edit-description")).clear();
	    driver.findElement(By.id("edit-description")).sendKeys(pitchDescription);
	    driver.findElement(By.id("edit-continue")).click();
	    driver.findElement(By.linkText("OK - I'm Ready!")).click();
	    
	    WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
	    fileInput.sendKeys(pathName);
	    driver.findElement(By.id("edit-continue")).click();
	    
	    driver.findElement(By.id("edit-submit")).click();
	    driver.findElement(By.linkText("Logout")).click();
	}
	
	public static void processPitchNewUser() throws BiffException, IOException
	{		  
		firstName = finalSheet.getCell(0, y).getContents();
		lastName = finalSheet.getCell(1, y).getContents();
		userNameEmail = finalSheet.getCell(2, y).getContents();
		schoolType = finalSheet.getCell(3, y).getContents();
		year = finalSheet.getCell(4,y).getContents();
		categoryName = finalSheet.getCell(4, y).getContents();
		pitchTitle = finalSheet.getCell(5, y).getContents();
		pitchDescription = finalSheet.getCell(6, y).getContents();
		pathName = finalSheet.getCell(6, y).getContents();
		
		System.out.println(userNameEmail);
		System.out.println(passwordName);
		System.out.println(categoryName);
		System.out.println(pitchTitle);
		System.out.println(pitchDescription);
		System.out.println(pathName);
		
		y++;
		
		driver.get("http://1000pitches.com");
		driver.findElement(By.linkText("Pitch Your Idea")).click();
		new Select(driver.findElement(By.id("edit-university"))).selectByVisibleText("University of Southern California");
		driver.findElement(By.id("edit-first-name")).sendKeys(firstName);
		driver.findElement(By.id("edit-last-name")).sendKeys(lastName);
		driver.findElement(By.id("appendedInput")).sendKeys(userNameEmail);
		driver.findElement(By.id("edit-umid")).sendKeys(firstName);
		
		new Select(driver.findElement(By.id("edit-college"))).selectByVisibleText(schoolType);
		new Select(driver.findElement(By.id("edit-grad-year"))).selectByVisibleText(year);
	    driver.findElement(By.id("edit-continue")).click();

	    driver.findElement(By.id("edit-title")).sendKeys(pitchTitle);
	    new Select(driver.findElement(By.id("edit-category"))).selectByVisibleText(categoryName);
	    driver.findElement(By.id("edit-station")).click();
	    driver.findElement(By.id("edit-description")).clear();
	    driver.findElement(By.id("edit-description")).sendKeys(pitchDescription);
	    driver.findElement(By.id("edit-continue")).click();
	    driver.findElement(By.linkText("OK - I'm Ready!")).click();
	    
	    WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
	    fileInput.sendKeys(pathName);
	    driver.findElement(By.id("edit-continue")).click();
	    driver.findElement(By.id("edit-submit")).click();
	    driver.findElement(By.linkText("Logout")).click();	
	}
}
