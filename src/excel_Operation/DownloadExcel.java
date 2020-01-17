package excel_Operation;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import excel_Operation.CommonMethod;

public class DownloadExcel {

	static String url ="http://url";
	public static WebDriver driver;
	private static final String FILE_DIR = "Excel Manipulation\\externalFiles\\downloadFiles";
	private static final String FILE_TEXT_EXT = ".xlsx";
	static CommonMethod com = new CommonMethod();
	
	public static void main(String[] args) throws InterruptedException {

		driver = com.setting_download_Option();
		driver.get(url);
		System.out.println("url launched");
		driver.manage().window().maximize(); 

		//Deleting Existing File
		com.deleteFile(FILE_DIR, FILE_TEXT_EXT);
		// Click to download 
		driver.findElement(By.xpath("//html[@attribute='value']")).click();
		
	    Thread.sleep(5000);
	    driver.close();
	}

}
