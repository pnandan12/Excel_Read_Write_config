package readDataXL;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXL {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File src = new File("./exceldata.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb =new  XSSFWorkbook(fis);
		   XSSFSheet sheet1 =  wb.getSheetAt(0);
		 String data0 = sheet1.getRow(0).getCell(0).getStringCellValue();
         System.out.println("Data from Excel is " + data0);
         String data1 = sheet1.getRow(0).getCell(1).getStringCellValue();
         System.out.println("Data from Excel is " + data1);
	wb.close();
	}

}
