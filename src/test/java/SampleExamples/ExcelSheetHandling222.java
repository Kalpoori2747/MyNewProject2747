package SampleExamples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class ExcelSheetHandling222 {

	public static WebDriver driver;
	public static void main(String[] args) throws InterruptedException, IOException
	{
		driver = new ChromeDriver();
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
	    System.out.println("Application Opened");
	    
	    
//	    WebElement popup=driver.findElement(By.xpath("//div[@class='glow-toaster-footer']"));
//	    if(popup.isDisplayed())
//	    {
//		    Thread.sleep(2000);
//	    	driver.findElement(By.xpath("(//span[@class='a-button-inner']/input)[1]")).click();
//		    System.out.println("Popup dismiss is clicked");
//	    }
	    Thread.sleep(10000);
	    WebElement dropdown=driver.findElement(By.id("searchDropdownBox"));
	    dropdown.click();
	    System.out.println("Dropdown is clicked");
	    Thread.sleep(2000);
	    
	    Select sc=new Select(dropdown);
	    sc.selectByVisibleText("Books");
	    System.out.println("Books is selected");
	    Thread.sleep(2000);
	    
	    WebElement search=driver.findElement(By.xpath("//input[@type='submit']"));
	    search.click();
	    System.out.println("Search symbol is clicked");
	    Thread.sleep(2000);
	    
	    WebElement books=driver.findElement(By.xpath("(//div[@class='left_nav browseBox']/ul/li[1]/a)[1]"));
		books.click();
		System.out.println("New Year New Books is clicked");
		Thread.sleep(2000);
		
		WebElement Kindleunlimitedcheckbox=driver.findElement(By.xpath("//label[@for='apb-browse-refinements-checkbox_0']/i"));
		Kindleunlimitedcheckbox.click();
	    System.out.println("Kindleunlimitedcheckbox is clicked");
		Thread.sleep(2000);

		WebElement sortby=driver.findElement(By.xpath("//span[@class='a-dropdown-label']"));
		sortby.click();
	    System.out.println("sortby is clicked");
		Thread.sleep(2000);
		
		WebElement list=driver.findElement(By.xpath("//div[@class='a-popover-inner']/ul/li[6]"));
		list.click();
	    System.out.println("Best sellers is selected from the sortby dropdown list");
		Thread.sleep(2000);
		
		List<WebElement> results = driver.findElements(By.xpath("//div[@data-component-type='s-search-result']"));
		int size=results.size();
		System.out.println("The number of products present in the list : "+size);
		System.out.println("****************************************************");
		
		
		XSSFWorkbook wb = new XSSFWorkbook();
	    XSSFSheet sh = wb.createSheet("Sheet1");
		//for(WebElement element:results
        for(int i=0;i<size;i++)
        {
        	//String product=element.getText();
		    String product=results.get(i).getText();
		    XSSFRow row = sh.getRow(i);
		    if (row == null)
		    {
		    	row = sh.createRow(i);
			    XSSFCell cell = row.createCell(0); 
			    cell.setCellValue(product);
		    }
		    FileOutputStream fos = new FileOutputStream("C:\\Users\\ADMIN\\Documents\\amazonProducts.xlsx");
		    wb.write(fos);
		    System.out.println(product);
		 }

	}

}
