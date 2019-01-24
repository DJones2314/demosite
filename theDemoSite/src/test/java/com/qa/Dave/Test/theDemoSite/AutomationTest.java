package com.qa.Dave.Test.theDemoSite;



import static org.junit.Assert.assertEquals;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.PageFactory;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import pages.AddUserPage;
import pages.Home;
import pages.login;
import utility.Constant;
import utility.ExcelUtils;

public class AutomationTest {

	WebDriver driver;
	
	static ExtentReports report;
	static ExtentTest test;
	static ExtentTest test2;
	static ExtentTest test3;
	
	

		@BeforeClass
		public static void preReport() {
		report = new ExtentReports(Constant.folderPath + Constant.reportFile , true);
		}
		@Before
		public void setup() {
		System.setProperty("webdriver.chrome.driver", Constant.driverLocation);
		driver = new ChromeDriver();
		test = report.startTest("demoSiteLogin");
		test2 = report.startTest("demoSiteLoginWithExcel");
		driver.get(Constant.siteLocation);
		driver.manage().window().maximize();
		} 
	
	@Test
	public void methodTest() {
		
		Home page = PageFactory.initElements(driver, Home.class);
		page.addUserPage(driver);
		test.log(LogStatus.INFO, "opening the add user page");
		AddUserPage page1 = PageFactory.initElements(driver, AddUserPage.class);
		page1.createUser("dave", "evad", driver);
		test.log(LogStatus.INFO, "adding the username and password");
		login page2 = PageFactory.initElements(driver, login.class);
		page2.loggingIn("dave", "evad", driver);
		test.log(LogStatus.INFO, "logging in with username and password");
		
		try {
				assertEquals("**Successful Login**", page2.check(driver)); 
				test.log(LogStatus.PASS, "Successfully logged on");
		} catch (Exception E) {
			test.log(LogStatus.FAIL, "Uanable to login");
		}
		}
	
	@Test
	public void excelTest() throws IOException, InterruptedException {

		FileInputStream file = new FileInputStream(Constant.pathTestData + Constant.fileTestData);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		ExcelUtils.setExcelFile(Constant.pathTestData + Constant.fileTestData, 0);

		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {

			test3 = report.startTest("writing results to the excel spreadsheet" + i);
			Cell username = sheet.getRow(i).getCell(0);
			Cell password = sheet.getRow(i).getCell(1);

			String user = username.getStringCellValue();
			String pass = password.getStringCellValue();
			
			Home page = PageFactory.initElements(driver, Home.class);
			page.addUserPage(driver);
			test2.log(LogStatus.INFO, "opening the add user page");
			test3.log(LogStatus.INFO, "opening the add user page");

			AddUserPage page1 = PageFactory.initElements(driver, AddUserPage.class);
			page1.createUser(user, pass, driver);
			test2.log(LogStatus.INFO, "adding the username and password");
			test3.log(LogStatus.INFO, "adding the username and password");
			
			login page2 = PageFactory.initElements(driver, login.class);
			page2.loggingIn(user, pass, driver);
			test2.log(LogStatus.INFO, "logging in with username and password");
			test3.log(LogStatus.INFO, "logging in with username and password");
			
			Thread.sleep(1000);
			
			try {
				assertEquals("**Successful Login**", page2.check(driver)); 
				test2.log(LogStatus.PASS, "Successfully logged on");
				test3.log(LogStatus.PASS, "Successfully logged on");
				ExcelUtils.setCellData("Pass", i, 2); 
		} catch (Exception E) {
			test2.log(LogStatus.FAIL, "Uanable to login");
			test3.log(LogStatus.FAIL, "Uanable to login");
			ExcelUtils.setCellData("Fail", i, 2);
		}
			
		}
	}
	

	@After
	public void tearDown() {
		driver.quit();
	}
	@AfterClass
	public static void tearReport() {
		report.endTest(test);
		report.flush();
	}
	
}
