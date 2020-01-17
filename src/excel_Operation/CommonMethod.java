package excel_Operation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.DirectoryNotEmptyException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import com.monitorjbl.xlsx.StreamingReader;

public class CommonMethod {
	ChromeDriver driver;
	String dir = "C:\\Users\\AC49605\\Downloads\\chrome";
	int rowCount = 0;
	

	public ChromeDriver setting_download_Option() {
		System.setProperty("webdriver.chrome.driver",
				"\\driver\\chromedriver.exe");
		Map<String, Object> prefs = new HashMap<String, Object>();

		// Use File.separator as it will work on any OS
		prefs.put("download.default_directory",
				System.getProperty("user.dir") + File.separator + "externalFiles" + File.separator + "downloadFiles");

		// Adding cpabilities to ChromeOptions
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", prefs);

		// Printing set download directory
		System.out.println(options.getExperimentalOption("prefs"));

		// Launching browser with desired capabilities
		driver = new ChromeDriver(options);

		return driver;
	}

		
	public void deleteFile(String folder, String ext) {

		GenericExtFilter filter = new GenericExtFilter(ext);
		File dir = new File(folder);

		// list out all the file name with .txt extension
		String[] list = dir.list(filter);

		if (list.length == 0)
			return;

		File fileDelete;

		for (String file : list) {
			String temp = new StringBuffer(folder).append(File.separator).append(file).toString();
			fileDelete = new File(temp);
			boolean isdeleted = fileDelete.delete();
			System.out.println("file : " + temp + " is deleted : " + isdeleted);
		}
	}

	public class GenericExtFilter implements FilenameFilter {

		private String ext;

		public GenericExtFilter(String ext) {
			this.ext = ext;
		}

		public boolean accept(File dir, String name) {
			return (name.endsWith(ext));
		}
	}

	public int getFileName(String folder, String Inventory_name) throws IOException, InterruptedException {
		File file = new File(folder);
		String[] fileList = file.list();
		String newFilename = Inventory_name + ".xlsx";
		System.out.println(fileList);
		System.out.println("Inventory_name: " + Inventory_name);
		System.out.println(newFilename);

		for (String name : fileList) {
			System.out.println(name);
			if (name.contains(newFilename)) {
				System.out.println("filename is same as downloaded file" + name);
				rowCount = ReadExcel(name, folder);
			} else if (Inventory_name.contains(" - ")) {
			rowCount = ReadExcel(name, folder);
				System.out.println("It contains name plus source as well");
			} else {
				System.out.println("Please chek again");
			}
		}
		return rowCount;
	}

	
	@SuppressWarnings("resource")
	public int ReadExcel(String newFilename, String folder) throws IOException, InterruptedException {

		Thread.sleep(2000);
		String backslash = "\\";
		String filepath = folder + backslash + newFilename;
		System.out.println(filepath);
		FileInputStream fis = new FileInputStream(filepath);
		ZipSecureFile.setMinInflateRatio(0.0d);
		Thread.sleep(10000);
	     XSSFWorkbook workbook = new XSSFWorkbook(fis);
		String sheet_name = workbook.getSheetName(0);

		XSSFSheet sheet = workbook.getSheet(sheet_name);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		System.out.println(rowCount);
		Row row = sheet.getRow(0);
		int colNum = row.getLastCellNum();
		System.out.println("Total Number of Columns in the excel is : " + colNum);
		int rowNum = sheet.getLastRowNum() + 1;
		System.out.println("Total Number of Rows in the excel including header is : " + rowNum);
		System.out.println("Total Number of Rows in the excel excluding header is : " + rowCount);
		workbook.close();
		
		return rowCount;
	}
}
