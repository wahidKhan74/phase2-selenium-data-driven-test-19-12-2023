package com.simplilearn.testng;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;

public class SignInTest {

	// step1: create webdriver reference
	String siteUrl = "https://accounts.simplilearn.com/user/login";
	String driverPath = "drivers/windows/geckodriver.exe";
	WebDriver driver;
	WebDriverWait wait;

	// create xlsx reference
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;

	@BeforeMethod
	public void beforeTest() {

		// step2: set system properties for selenium dirver
		System.setProperty("webdriver.geckodriver.driver", driverPath);

		// step3: instantiate selenium webdriver
		driver = new FirefoxDriver();

		// step4: launch browser
		driver.get(siteUrl);
		driver.manage().window().maximize();

		wait = new WebDriverWait(driver, Duration.ofSeconds(50));
	}

	@AfterMethod
	public void afterTest() {
		driver.quit();
	}

	@Test
	public void signInTest() {

		try {
			// Import xlsx sheet
			File src = new File("testdata\\testdata-login.xlsx");

			// Load the file as FileInputStream.
			FileInputStream fileInput = new FileInputStream(src);

			// Load the workbook
			workbook = new XSSFWorkbook(fileInput);

			// Load the sheet in which data is stored. (0)
			sheet = workbook.getSheetAt(0);

			for (int row = 1; row < sheet.getLastRowNum(); row++) {
				// import data from cell : username
				cell = sheet.getRow(row).getCell(1);
				cell.setCellType(CellType.STRING);
				driver.findElement(By.name("user_login")).sendKeys(cell.getStringCellValue());

				// import data from cell : password
				cell = sheet.getRow(row).getCell(2);
				cell.setCellType(CellType.STRING);
				driver.findElement(By.name("user_pwd")).sendKeys(cell.getStringCellValue());

				driver.findElement(By.name("btn_login")).submit();

				driver.findElement(By.name("user_login")).clear();

				Thread.sleep(400);

				// Write data in the excel.
				FileOutputStream foutput=new FileOutputStream(src);
				
				// Specify the message needs to be written.
				String message = "Faliure Login Test";
				
				//Create cell where data needs to be written.
				sheet.getRow(row).createCell(3).setCellValue(message);
				
				// Specify the file in which data needs to be written.
				FileOutputStream fileOutput =new FileOutputStream(src);
				
				//write content
				workbook.write(fileOutput);
				
				  // close the file
				fileOutput.close();
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

	}

}
