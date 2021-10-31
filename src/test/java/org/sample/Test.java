package org.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Test {
	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\GOWTHAM\\OneDrive\\Documents\\gowtham.xlsx");
		
		 Workbook w=new XSSFWorkbook();
	 WebDriverManager.chromedriver().setup();
	 WebDriver driver=new ChromeDriver();
	 driver.get("https://www.seleniumeasy.com/test/table-search-filter-demo.html");
	 WebElement table = driver.findElement(By.tagName("table"));
	 WebElement heading = table.findElement(By.tagName("thead"));
	 WebElement headingrow = heading.findElement(By.tagName("tr"));
	 List<WebElement> headings = headingrow.findElements(By.tagName("th"));
	 Sheet sheet = w.createSheet("abcd");
		Row row = sheet.createRow(0);
	 for (int i = 0; i < headings.size(); i++) {
		 Cell cell = row.createCell(i);
		 WebElement headingElement = headings.get(i);
		 String data = headingElement.getText();
		 cell.setCellValue(data);
		
	}
	 WebElement body = table.findElement(By.tagName("tbody"));
	 List<WebElement> rows = body.findElements(By.tagName("tr"));
	 for (int i = 0; i < rows.size(); i++) {
		 WebElement individualrow = rows.get(i);
		  List<WebElement> datas =individualrow .findElements(By.tagName("td"));
		 
		  
		 Row row2 = sheet.createRow(i);
		for (int j = 0; j < datas.size(); j++) {
			Cell cell = row2.createCell(j);
			WebElement data = datas.get(j);
			String text = data.getText();
			cell.setCellValue(text);	
		}
	}
	  FileOutputStream str=new FileOutputStream(f);
	 w.write(str);
	 driver.quit();
	}

}
