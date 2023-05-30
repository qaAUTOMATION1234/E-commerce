package COM.TESTCASES;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class PropertyfileExample {
	public  WebDriver driver;
@Test
	
	public void ReadData() throws IOException
	{
		//Step 1 we give the location for the property file
		FileReader file = new FileReader("./InputTestData/TestData.properties");
		
		//Step 2 create an object for Properties class
		
		Properties prop = new Properties();
		
		//Step 3 load the file into property
		
		prop.load(file);
		
		System.out.println(prop.getProperty("userid"));
		
		String Name = prop.getProperty("Name");
		
		System.out.println(Name);
	}
@Test

public void WriteData() throws Exception
{
	WebDriverManager.chromedriver().setup();
	driver = new ChromeDriver();
	//WebDriverManager.chromedriver().setup();
	
	driver.manage().window().maximize();
	driver.get("https://automationexercise.com/");
	
	String url = driver.getCurrentUrl();
	String title = driver.getTitle();
	
	//Step 1 we give the location for the property file
	String file ="./OutputTestData/OutputData.properties";
			
			//Step 2 create an object for Properties class
			
	Properties prop = new Properties();
	try {
		//Creating output stream to write the file
		
		OutputStream os = new FileOutputStream(file);
		
		//Write the data to a property file
		
		prop.setProperty("URL", url);
		prop.setProperty("Title", title);
		
		prop.store(os, "Data Saved Page Title and URL");
	
		//Close the output stream
		
		os.close();
		
	
	}
	
	catch(Exception e)
	{
		System.out.println(e);
	}
	//Click on the menu links from the property file
			
				//Step 1 we give the location for the property file
					FileReader file1 = new FileReader("./InputTestData/MenuData.properties");
					
					//Step 2 create an object for Properties class
					
					Properties prop1 = new Properties();
					
					//Step 3 load the file into property
					
					prop1.load(file1);
					
					String WomenLocator = prop1.getProperty("WomenLocator");
					String MenLocator = prop1.getProperty("MenLocator");
					
					WebElement WomenMenu = driver.findElement(By.xpath(WomenLocator));
					WebElement MenMenu = driver.findElement(By.xpath(MenLocator));
					
					WomenMenu.click();
					MenMenu.click();
					
					
			
			
			
			driver.quit();
	driver.quit();
}
}
