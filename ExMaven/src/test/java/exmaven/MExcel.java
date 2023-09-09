package exmaven;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.File;
import java.io.FileInputStream;
import org.openqa.selenium.support.ui.Select;




public class MExcel {

	public static WebElement emlid;
	public static WebElement pwd;
	public static WebElement repwd;
	public static WebElement alteml;
	public static WebElement mob;
	public static WebElement dd;
	public static WebElement mm;
	public static WebElement yyyy;
	public static WebElement gender;
	public static WebElement cntry;
	public static WebElement city;
	private static XSSFWorkbook wb;
	 public  static WebDriver d;

		public static void main(String[] args) throws IOException
		{
			System.setProperty("webdriver.chrome.driver","C:\\Program Files\\Seleniumtrupti\\chromedriver-win64.exe");
			WebDriverManager.chromedriver().setup();
			 ChromeOptions opt = new ChromeOptions();
			 opt.addArguments("--remote-allow-origins=*");
			 opt.addArguments("--disable-notification");  
			 d = new ChromeDriver(opt);
			 
			d =new ChromeDriver();
			d.manage().window().maximize();
			
			d.get("https://www.rediff.com/");
			
			WebElement creacc = d.findElement(By.linkText("Create Account"));
			
			creacc.click(); 
			
			WebElement fn = d.findElement(By.xpath("//*[@id=\'tblcrtac\']/tbody/tr[3]/td[3]/input"));
			fn.click();
			
			File file = new File("C:\\Users\\Trupti\\ExMaven\\src\\test\\resources\\Data.xlsx");
			
			FileInputStream fis = new FileInputStream(file);
			
			wb = new XSSFWorkbook(fis);
			
			XSSFSheet sheet = wb.getSheet("Sheet1");
			
			int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
			
			for (int i = 1; i<=rowCount; i++)
			{
				int cellCount = sheet.getRow(i).getLastCellNum();
				
				for(int j =0; j<cellCount; j++)
				{
					
					String dt = sheet.getRow(i).getCell(j).getStringCellValue();
					
						if(j==0)
					{
						fn.sendKeys(dt);
					}
					if(j==1)
					{
						emlid = d.findElement(By.xpath("//*[@id=\'tblcrtac\']/tbody/tr[7]/td[3]/input[1]"));
						emlid.click();
						emlid.sendKeys(dt);
					}
					if(j==2)
					{
						pwd = d.findElement(By.xpath("//*[@id=\"newpasswd\"]"));
						pwd.click();
						pwd.sendKeys(dt);
					}
					if(j==3)
					{
						repwd = d.findElement(By.xpath("//*[@id=\"newpasswd1\"]"));
						repwd.click();
						repwd.sendKeys(dt);
					}
					if(j==4)
					{
						alteml = d.findElement(By.xpath("//*[@id=\"div_altemail\"]/table/tbody/tr[1]/td[3]/input"));
						alteml.click();
						alteml.sendKeys(dt);
					}
					if(j==5)
					{
						mob = d.findElement(By.xpath("//*[@id=\'mobno\']"));
						mob.click();
						mob.sendKeys(dt);
						
					}
					if(j==6)
					{
						dd = d.findElement(By.cssSelector("#tblcrtac > tbody > tr:nth-child(23) > td:nth-child(3) > select:nth-child(1)"));
						dd.click();
						
						Select date = new Select(dd);
						
						date.selectByVisibleText(dt);
						
					}
					if(j==7)
					{
						mm = d.findElement(By.xpath("//*[@id=\'tblcrtac\']/tbody/tr[22]/td[3]/select[2]"));
						
						Select month = new Select (mm);
						
						month.selectByVisibleText(dt);
					}
					if(j==8)
					{
						yyyy = d.findElement(By.cssSelector("#tblcrtac > tbody > tr:nth-child(23) > td:nth-child(3) > select:nth-child(3)"));
						
						Select year = new Select(yyyy);
						
						year.selectByVisibleText(dt);
					}
					if(j==9)
					{
						gender = d.findElement(By.xpath("//*[@id=\'tblcrtac\']/tbody/tr[24]/td[3]/input[2]"));
						gender.click();
						
					}
					if(j==10)
					{
						cntry = d.findElement(By.xpath("//*[@id=\'country\']"));
						
						Select country = new Select (cntry);
						
						country.selectByVisibleText(dt);
					}
					if(j==11)
					{
						city = d.findElement(By.xpath("//*[@id=\'div_city\']/table/tbody/tr[1]/td[3]/select"));
						
						Select ct = new Select (city);
						 
						ct.selectByVisibleText(dt);
						
					}
					
					
				}
				
				
			}
			
			

		}
	}


