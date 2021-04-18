package readDataXL;

import libr.ExcelDataConfig;

public class ReadXLdataFromDataConfig {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
 ExcelDataConfig excel= new ExcelDataConfig("./exceldata.xlsx");
 
	System.out.println(excel.getData(0,0,0));
	System.out.println(excel.getData(1,0,1));
	
	}

}
