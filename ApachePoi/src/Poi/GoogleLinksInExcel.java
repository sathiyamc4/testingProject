package Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class GoogleLinksInExcel 
{
	public static void main(String args[]) throws IOException
	{
		int j=0;
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\SATHIYA\\eclipse-workspace\\Selenium\\Driver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.google.co.in/?gws_rd=ssl");
		FileOutputStream fout = new FileOutputStream("C:\\Users\\SATHIYA\\Desktop\\sample.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Sheet2");
		List<WebElement> lists = driver.findElements(By.tagName("a"));
		System.out.println("No. of links:   "+lists.size());
		for(WebElement list: lists)
		{
			System.out.println(list.getAttribute("href"));
			sheet.createRow(j).createCell(0).setCellValue(list.getAttribute("href"));
			j++;
		}
		wb.write(fout);
		driver.close();
	}
}
