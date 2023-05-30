package COM.TESTCASES;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;
public class Writetoexceldata {
	public WebDriver driver;
	@Test
	public void GetURL() throws EncryptedDocumentException, IOException
	{
		WebDriverManager.chromedriver().setup();
		
		driver = new ChromeDriver();
		
		FileInputStream fout = new FileInputStream("./OutputTestData/data.xlsx");
		
		driver.manage().window().maximize();
		
		
		driver.get("https://automationexercise.com/");
		String url = driver.getCurrentUrl();
		String title = driver.getTitle();
		
		Workbook wb = WorkbookFactory.create(fout);
		
		Sheet sh = wb.getSheet("PageDetails");
		
		Row rw = sh.createRow(0);
		Row rw1 = sh.createRow(1);
		
		
		Cell cell = rw.createCell(0);
		Cell cell1 = rw1.createCell(0);
		
		cell.setCellValue(url);
		cell1.setCellValue(title);
		int rowNum = 0;
		List<WebElement> links=driver.findElements(By.tagName("a"));
		for(WebElement link : links)
		{
			Row row = sh.createRow(rowNum++);
			Cell cell2 = row.createCell(1);
			cell2.setCellValue(link.getAttribute("href"));
			
		}
		
		//Final step is to wirte the values into the Excel file
		
		FileOutputStream fout1 = new FileOutputStream("./OutputTestData/Output.xlsx");
		
		wb.write(fout1);
		
		driver.quit();
	}

}
