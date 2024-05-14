package SeleniumExamples;

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

public class ExcelSheet2747 {

	public static WebDriver driver;
	public static void main(String[] args) throws IOException {
		
		driver=new ChromeDriver();
		
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get("https://www.google.com/");
		driver.manage().window().maximize();
		
		List<WebElement>links=driver.findElements(By.tagName("a"));
		
		int size=links.size();
		
		XSSFWorkbook wb=new XSSFWorkbook();
		XSSFSheet sheet=wb.createSheet("Sheet1");
		
		for(int i=0;i<size;i++) {
			
			String values=links.get(i).getText();
			
			XSSFRow row=sheet.getRow(i);
			
			if(row==null) {
				row=sheet.createRow(i);
				
				XSSFCell cell=row.createCell(0);
				cell.setCellValue(values);
			}
			
			FileOutputStream fo=new FileOutputStream("D:\\SeleniumPracticeWork\\LatestOne11\\TestData\\Googlelinks.xlsx");
			
			wb.write(fo);
			System.out.println(values);
		}
		

	}

}
