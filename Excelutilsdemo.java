package utils;

public class Excelutilsdemo {
  public static void main(String[] args) {
	  
		String projectpath = System.getProperty("user.dir");
	  excelutils excel = new excelutils(projectpath+"/excel/data.xlsx","sheet1");
	  
	  excelutils.getrowcount();
	  excelutils.getcolcount();
	  excelutils.getcelldataString(0, 0);
	//  excel.getcelldataNumber(1, 1);
	  
  }
 
}

