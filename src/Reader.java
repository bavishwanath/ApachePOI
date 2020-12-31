import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reader {

	static XSSFSheet sheet;

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		String path = "D://Test Data//healthdata.xlsx";
		FileInputStream fis = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		System.out.println(sheets);
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("data1")) {
				sheet = workbook.getSheetAt(i);
			}
		}
		
		// scan the headers
		
	}

}
