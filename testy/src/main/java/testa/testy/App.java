package testa.testy;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.WatchKey;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App 
{
    public static void main( String[] args ) throws  IOException
    {
    	// Making a connection to the excel file
    	
    	//FileInputStream fs = new  FileInputStream("C:\\Users\\ilabadmin\\Desktop\\Book1.xlsx");
    	FileOutputStream fs = new  FileOutputStream("C:\\Users\\ilabadmin\\Desktop\\Book1.xlsx"); // Write on excel file
    	
    	XSSFWorkbook wb = new XSSFWorkbook();
    	//create a workbook object
    	
    	/**XSSFWorkbook wb = new XSSFWorkbook(fs);
    	XSSFSheet s1 =  wb.getSheet("Sheet1");
    	XSSFRow  r1 = s1.getRow(0);
    	XSSFCell c1 = r1.getCell(1);
    	System.out.println(c1.getStringCellValue());  // Display data at cell 1**/
    	
    	//System.out.println("Display number of rows" + " " +s1.getPhysicalNumberOfRows());  // Display number of rows

    	//System.out.println("Displayindex of last row" + " " +s1.getLastRowNum());  // Display number of rows
    	XSSFSheet s1 = wb.createSheet("Result");
    	XSSFRow  r1 = s1.createRow(0);
    	XSSFCell c1 = r1.createCell(1);
    	c1.setCellValue("hello");
    	wb.write(fs);
    	wb.close();
    	
    	
    }
}
