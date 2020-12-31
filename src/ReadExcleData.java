import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcleData {
	String filepath = "D:\\Test Data\\healthdata.xlsx";

	XSSFWorkbook workbook = null;

	public void readExcel(String filePath, String fileName, String sheetName) throws IOException {
		File file = new File(filePath + "\\" + fileName);
		FileInputStream fis = new FileInputStream(file);
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		if (fileExtensionName.equals(".xlsx")) {
			workbook = new XSSFWorkbook(fis);
		} else if (fileExtensionName.equals(".xls")) {
			workbook = new XSSFWorkbook(fis);
		}
		Sheet sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for (int i = 0; i < rowCount + 1; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				System.out.print(row.getCell(j).getStringCellValue() + "|| ");

			}
			System.out.println();

		}

	}

	public static void main(String args[]) throws IOException {
		ReadExcleData readExcel = new ReadExcleData();
		String filePath = "D:\\Test Data";
		String fileName = "healthdata.xlsx";
		String sheetName = "health";
		readExcel.readExcel(filePath, fileName, sheetName);
	}

}
