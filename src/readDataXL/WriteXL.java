package readDataXL;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteXL {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File src = new File("./exceldata.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb =new  XSSFWorkbook(fis);
		   XSSFSheet sheet1 =  wb.getSheetAt(0);
		 sheet1.getRow(0).createCell(2).setCellValue("Pass");
		 sheet1.getRow(1).createCell(2).setCellValue("Fail");
		 
		 FileOutputStream fout= new FileOutputStream(src);
		 wb.write(fout);
		 
		 
	wb.close();
	}

}
