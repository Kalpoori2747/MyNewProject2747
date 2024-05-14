package SeleniumExamples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWritingExamples {

	public static void main(String[] args) throws IOException {
		//it is converted excel sheet into writing mode
		FileOutputStream fo=new FileOutputStream("D:\\SeleniumPracticeWork\\LatestOne11\\TestData\\ExcelSheetWriting.xlsx");

		XSSFWorkbook wb=new XSSFWorkbook();
		XSSFSheet sheet=wb.createSheet();
		//we need to provide the data at the time running the testcase
		Scanner sc=new Scanner(System.in);
		
		for(int r=0;r<=3;r++) {//it is represents the rows
			//create the rows
			XSSFRow row=sheet.createRow(r);
			
			for(int c=0;c<3;c++) { //it is represents the columns
				
				//System.out.println("enter a value");
				
				String value=sc.next();//it is only accepted string values
			   row.createCell(c).setCellValue(value);
			}
		}
		
		wb.write(fo);
		wb.close();
		fo.close();
		System.out.println("it's done");
		
	}

}
