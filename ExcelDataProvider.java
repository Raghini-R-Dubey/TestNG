package utils;
import org.junit.Before;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider {
	
	 WebDriver driver=new ChromeDriver();
	
	@Before
	public void setuptest() {
		String projectpath = System.getProperty("user.dir");
		System.setProperty("webdriver.chrome.driver", projectpath+"/excel/data.xlsx");
	//	driver = new ChromeDriver();
	}
	@Test(dataProvider ="test1data")
	public void test1(String username, String password) throws InterruptedException{
		System.out.println(username+" | "+password);
		
		driver.get("https://opensource-demo.orangehrmlive.com/");
		driver.findElement(By.id("txtUsername")).sendKeys(username);
		driver.findElement(By.id("txtPassword")).sendKeys(password);
		Thread.sleep(2000);
	}
	
	@DataProvider(name= "test1data")
	public Object[][] getdata() {
		
		String excelpath="/home/am-pc-45/eclipse-workspace/poiapache/excel/data.xlsx";
		 Object data[][] = testdata(excelpath, "sheet1");
		return data;
		
		
		
	}
   public  Object[][] testdata(String excelpath, String sheetName) {
	   
	   
		String projectpath = System.getProperty("user.dir");
	   excelutils excel = new excelutils(projectpath+"/excel/data.xlsx","sheet1");
	   
	   int rowcount = excel.getrowcount();
	  int colCount =  excel.getcolcount();
	  
	  Object data[][] = new Object[rowcount-1][colCount];
	  
	  for(int i=1; i<rowcount; i++) {
		  for(int j=0; j<colCount; j++){
			  
			  String celldata = excel.getcelldataString(i, j);
			 // System.out.print(celldata+ " | ");
			  data[i-1][j]=	celldata;
			  }
		  //System.out.println( );
	  }
	return data;
   } 
  
}

