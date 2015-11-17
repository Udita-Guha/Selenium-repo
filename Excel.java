import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xalan.lib.sql.ObjectArray;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


public class Excel {
	
	public WebDriver driver;
	String baseUrl;
	String Filepath="C://Users//Admin//Documents//excel1.xlsx";
	String value;
	@BeforeTest
	public void GmailURL(){
		driver=new FirefoxDriver();
		driver.manage().window().maximize();
		baseUrl="https://www.gmail.com";
		driver.get(baseUrl);
		}
	
	@Test(dataProvider="Exceldata")
	public void ExcelRead(){
		driver.findElement(By.id("Email")).sendKeys("value");
		driver.findElement(By.id("next")).click();
		driver.findElement(By.id("Passwd")).sendKeys("value");
		driver.findElement(By.id("signIn")).click();
		
		WebElement actual=driver.findElement(By.xpath("//div[@class='form-panel second']/descendant::span[@id='errormsg_0_Passwd']"));
		String actualmsg=actual.getText();
		String requiredMessage = "Please correct the marked field(s) below.";
		Assert.assertEquals(requiredMessage, actualmsg);
		
		
		
	}
	
	@DataProvider
	public Object[][] ExcelData() throws IOException{
		Object[][] ExcelData=null;
		FileInputStream file=new FileInputStream(Filepath);
		XSSFWorkbook wrkbk=new XSSFWorkbook(file);
		XSSFSheet sheet=wrkbk.getSheet("login");
		int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum()+1;
		Row r=sheet.getRow(0);
		int cellcount=r.getLastCellNum();
		ExcelData= new String[rowcount][cellcount];
			for(int i=0;i<=rowcount;i++){
			for(int j=0;j<=cellcount;j++){
				Cell cell=r.createCell(j);
				ExcelData[i][j]=cell.getRichStringCellValue();
			}		
		}
		
		
		return ExcelData;
		
	}
	
	@AfterTest
	public void end(){
		driver.close();
	}

}
