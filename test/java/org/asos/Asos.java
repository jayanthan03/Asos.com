package org.asos;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Asos {
	
	public static String getData(int rowNo,int cellNo) throws Throwable {
	String v=null;
	File loc =new File("C:\\Users\\Jayanthan\\eclipse-workspace\\org.asos\\TestData\\excelread.xlsx");
	FileInputStream stream=new FileInputStream(loc);
	Workbook w= new XSSFWorkbook(stream);
	Sheet s = w.getSheet("Sheet1");
	Row r = s.getRow(rowNo);
	Cell c = r.getCell(cellNo);
	int type = c.getCellType();
	if(type==1) {
		 v = c.getStringCellValue();
	}
	else if(type==0) {
		if(DateUtil.isCellDateFormatted(c))
		{
			Date dateCellValue = c.getDateCellValue();
			SimpleDateFormat sim=new SimpleDateFormat("dd-MM-yy");
			v = sim.format(dateCellValue);
		}
		else {
			double numericCellValue = c.getNumericCellValue();
			long l=(long)numericCellValue;
			v = String.valueOf(l);
		}
	}
	
return v;
	
}
    

	public static void main(String[] args) throws Throwable {
     System.setProperty("webdriver.chrome.driver", "C:\\Users\\Jayanthan\\eclipse-workspace\\org.asos\\driver\\chromedriver.exe");
	WebDriver driver=new ChromeDriver();
	driver.manage().window().maximize();
	driver.navigate().to("https://www.asos.com/");
	WebElement e1 = driver.findElement(By.xpath("(//a[text()='MEN'])[1]"));
	JavascriptExecutor js=(JavascriptExecutor)driver;
	js.executeScript("arguments[0].click()", e1);
	WebElement e2 = driver.findElement(By.xpath("(//span[text()='Shoes'])[2]"));
	Actions ac=new Actions(driver);
	ac.moveToElement(e2).perform();
	WebElement e3 = driver.findElement(By.xpath("(//a[text()='adidas'])[2]"));
	js.executeScript("arguments[0].click()", e3);
	driver.findElement(By.xpath("//p[text()='adidas Originals POD trainers in black']")).click();
	WebElement e4 = driver.findElement(By.xpath("(//select[@data-id='sizeSelect'])[1]"));
	Select s=new Select(e4);
	s.selectByVisibleText("UK 7 - EU 40.5 - US 7.5");
	driver.findElement(By.xpath("(//span[text()='Add to bag'])[1]")).click();
	Thread.sleep(2000);
	WebElement e5 = driver.findElement(By.xpath("//span[@class='_1z5n7CN']"));
	ac.moveToElement(e5).perform();
	TakesScreenshot tk=(TakesScreenshot)driver;
	File src=tk.getScreenshotAs(OutputType.FILE);
	File desc=new File("C:\\Users\\Jayanthan\\eclipse-workspace\\org.asos\\ScreenShot\\ScreenShot1.jpeg");
	FileUtils.copyFile(src, desc);
	/*WebElement e6 = driver.findElement(By.xpath("//span[text()='Checkout']"));
	js.executeScript("arguments[0].click()", e6);
	driver.findElement(By.xpath("//input[@name='Username']")).sendKeys(getData(0,0));
	driver.findElement(By.xpath("//input[@aria-labelledby='PasswordLabel']")).sendKeys(getData(1,0));
	driver.findElement(By.xpath("//input[@type='submit']")).click();
	//driver.quit();*/
	driver.findElement(By.xpath("//span[text()='View Bag']")).click();
	Thread.sleep(2000);
	WebElement e7 = driver.findElement(By.xpath("//a[text()='adidas Originals POD trainers in black']"));
	String itemname = e7.getText();
	WebElement e8 = driver.findElement(By.xpath("//span[@class='bag-item-price bag-item-price--current bag-item-price--markedDown']"));
	String price = e8.getText();
	//WebElement e9 = driver.findElement(By.xpath("(//p[@class='bag-total-title-holder bag-total-title-holder--subtotal'])[2]"));
	//String text = e9.getText();
	File loc1=new File("C:\\Users\\Jayanthan\\eclipse-workspace\\org.asos\\TestData\\excelwrite.xlsx");
	Workbook w1=new XSSFWorkbook();
	Sheet s3=w1.createSheet("order summary");
	Row r=s3.createRow(0);
	Cell c = r.createCell(0);
	c.setCellValue("item name");
	Cell c1 = r.createCell(1);
	c1.setCellValue("Price");
	Row r1 = s3.createRow(1);
	Cell c2 = r1.createCell(0);
	c2.setCellValue(itemname);
	Cell c3 = r1.createCell(1);
	c3.setCellValue(price);
	FileOutputStream o=new FileOutputStream(loc1);
	w1.write(o);
	
	}

}
