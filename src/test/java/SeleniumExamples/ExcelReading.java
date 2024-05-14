package SeleniumExamples;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReading {

	public static void main(String[] args) throws IOException {
		
		//it is coverted into reading mode
		FileInputStream fis=new FileInputStream("D:\\SeleniumPracticeWork\\LatestOne11\\TestData\\ExcelReading1747.xlsx");

		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Sheet1");
		
		//identify the rows and columns
		int rows=sheet.getLastRowNum();
		int columns=sheet.getRow(rows).getLastCellNum();
		
		System.out.println(rows);
		System.out.println(columns);
		
		for(int i=1;i<=rows;i++) {//represents the rows
			//capture the rows
			XSSFRow row=sheet.getRow(i);
			
			for(int c=0;c<columns;c++) {  //represents columns
		       
				String values=row.getCell(c).toString();
				
				System.out.print(values);
			
			}
			System.out.println();
	}

	}
}
