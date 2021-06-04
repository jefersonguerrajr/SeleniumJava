package TestandoSelenium.Teste;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		// path to chromedriver.exe
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();

		try {
			driver.get("https://www.w3schools.com/html/html_tables.asp");


			int colunas = 0;

			for (int i = 1; i < 100; i++) {

				try {
					String xpath = ("//*[@id=\"customers\"]/tbody/tr[1]/th[" + i + "]");
					driver.findElement(By.xpath(xpath)).getText();

				} catch (NoSuchElementException ex) {
					colunas = i - 1;
					break;

				}

			}

			System.out.println(colunas);

			int linhas = 0;

			for (int i = 2; i < 100; i++) {

				try {
					String xpath = ("//*[@id=\"customers\"]/tbody/tr[" + i + "]/td[1]");
					driver.findElement(By.xpath(xpath)).getText();

				} catch (NoSuchElementException noSuchElementException) {
					linhas = i - 1;
					break;

				}

			}

			System.out.println(linhas);

			// workbook object
			XSSFWorkbook workbook = new XSSFWorkbook();

			// spreadsheet object
			XSSFSheet spreadsheet = workbook.createSheet(" Student Data ");

			// creating a row object
			XSSFRow row;

			// This data needs to be written (Object[])
			Map<String, List> studentData = new TreeMap<String, List>();
			
			
			//*[@id="customers"]/tbody/tr[2]/td[1]
			//*[@id="customers"]/tbody/tr[2]/td[2]
			//*[@id="customers"]/tbody/tr[2]/td[3]
		
			int id = 1;
			
			for (int i = 2; i <= linhas; i++) {
				List<String> linha = new ArrayList<>();
				
				for (int j = 1; j <= colunas; j++) {
					String xpath = ("//*[@id=\"customers\"]/tbody/tr[" + i + "]/td["+j+"]");
					String result = driver.findElement(By.xpath(xpath)).getText();
					linha.add(result);
					
				}
				
				studentData.put(String.valueOf(id), linha);
				id++;

							
				
			}
			
			System.out.println(studentData);
			
			Set<String> keyid = studentData.keySet();

			int rowid = 0;

			// writing the data into the sheets...

			for (String key : keyid) {

				row = spreadsheet.createRow(rowid++);
				List objectArr = studentData.get(key);
				int cellid = 0;

				for (Object obj : objectArr) {
					Cell cell = row.createCell(cellid++);
					cell.setCellValue((String) obj);
				}
			}

			// .xlsx is the format for Excel Sheets...
			// writing the workbook into the file...
			FileOutputStream out;
			out = new FileOutputStream(new File("C:/Planilhas/teste.xlsx"));
			workbook.write(out);
			out.close();

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			driver.quit();

		}
	}
}
