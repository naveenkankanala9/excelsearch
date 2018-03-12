package com.fedex.excelsearch;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

@SuppressWarnings({ "rawtypes" })
public class MaintenanceOnsiteSupportExpiry {

	WebDriver driver;

	public void invokeBrowser() {
		try {
			FileInputStream inputSteam = new FileInputStream(
					new File("C:\\Users\\navee\\Downloads\\GoogleSEarch.xlsx"));
			XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(inputSteam);
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			for (int counter = 0; counter < sheet.getLastRowNum(); counter += 20) {
				System.setProperty("webdriver.chrome.driver",
						"C:\\Users\\navee\\Downloads\\selenium-java-3.10.0\\chromedriver_win32\\chromedriver.exe");
				driver = new ChromeDriver();
				driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
				// driver.manage().timeouts().pageLoadTimeout(15, TimeUnit.SECONDS);
				driver.get("https://support.hpe.com/hpsc/wc/public/home");
				driver.findElement(By.className("hpui-secondary-text1")).click();
	
				Iterator<Row> rowIterator = sheet.iterator();		
				// Traversing over each row of XLSX file
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					if (row.getRowNum() != 0) // skip title row
					{
						Iterator cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = (Cell) cellIterator.next();
							driver.findElement(By.id("serialNumber" + counter)).sendKeys(cell.getStringCellValue());
							counter++;
						}
					}
					if(counter==20){
						break;
					}
				}
				driver.findElement(By.className("hpui-primary-button")).click();
				
				Iterator<Row> rowIterator1 = sheet.iterator();
				while (rowIterator1.hasNext()) {
					Row row = rowIterator1.next();
					if (row.getRowNum() != 0) // skip title row
					{
						Iterator cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = (Cell) cellIterator.next();
							row.createCell(counter);
							if (driver.findElements(By.xpath("//*[@id='generate_table_" + cell.getStringCellValue()
									+ "']/table/tbody/tr[6]/td[5]")).size() > 0) {
								row.createCell(1)
										.setCellValue(driver
												.findElement(By.xpath("//*[@id='generate_table_"
														+ cell.getStringCellValue() + "']/table/tbody/tr[6]/td[5]"))
												.getText());
							} else if (driver.findElements((By.xpath("//*[@id='generate_table_"
									+ cell.getStringCellValue() + "']/table/tbody/tr[1]/td[5]"))).size() > 0) {
								row.createCell(1)
										.setCellValue(driver
												.findElement(By.xpath("//*[@id='generate_table_"
														+ cell.getStringCellValue() + "']/table/tbody/tr[1]/td[5]"))
												.getText());
							} else {
								row.createCell(1).setCellValue("null");
							}
						}
					}
				}
				if(counter==20){
					break;
				}
			}
			FileOutputStream outputStream = new FileOutputStream("C:\\Users\\navee\\Downloads\\GoogleSEarch.xlsx");
			workbook.write(outputStream);
			outputStream.close();

			// driver.quit();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		MaintenanceOnsiteSupportExpiry obj1 = new MaintenanceOnsiteSupportExpiry();
		obj1.invokeBrowser();

	}

}
