package COM.TESTCASES;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;


public class ReadDataFromExcel {
@Test
	
	public void ReadDataFromExcelFile() throws EncryptedDocumentException, IOException
	{
		
		//Step 1 Read the excel file from the location
		
		FileInputStream fs = new FileInputStream("./InputTestData/Book1.xlsx");
		
		//Step 2 assing the file to a Workbook class
		
		Workbook wb = WorkbookFactory.create(fs);
		
		//Step 3 Read the work sheet
		
		Sheet sh = wb.getSheet("Names");
		
		//Step 4 Read the row value
		
		Row rw = sh.getRow(1);
		
		//Step 5 To read the col value
		
		Cell celdata = rw.getCell(0);
		
		System.out.println("The value from the Excel sheet is" +celdata.getStringCellValue());
		
	}
}
