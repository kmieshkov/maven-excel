import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class dataDriven {

	public static String sheetName = "TestCases";

	public static void main(String[] args) throws IOException {

		parseTestcases("/testcases.xlsx");
		parseTestData("/testdata.xlsx");

	}

	private static void parseTestData(String fileName) throws IOException {
		FileInputStream fileInputStream = new FileInputStream("/Users/kmieshkov/Projects/IdeaProjects/maven-excel" + fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		
	}

	public static void parseTestcases(String fileName) throws IOException {
		FileInputStream fileInputStream = new FileInputStream("/Users/kmieshkov/Projects/IdeaProjects/maven-excel" + fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

		ArrayList<String> columnNames = new ArrayList<>();
		HashMap<String, ArrayList<String>> data = new HashMap<>();
		
		XSSFSheet sheet = workbook.getSheet(sheetName);
		Iterator<Row> rows = sheet.iterator();

		// iterator for each row in a sheet
		for (int j = 0; rows.hasNext(); j++) {
			Row row = rows.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			ArrayList<String> array = new ArrayList<>();

			// iterator for each cell in a row
			while(cellIterator.hasNext()) {
				Cell cellValue = cellIterator.next();
				if (j == 0) {
					columnNames.add(cellValue.getStringCellValue());
					data.put(cellValue.getStringCellValue(), new ArrayList<>());
				} else {
					array.add(cellValue.getStringCellValue());
				}
			}

			// save data in format "ID" => ["TC-01", "TC-02" ...]
			// save data in format "Title" => ["....", "...." ...]
			for (int k = 0; k < array.size(); k++) {
				data.get(columnNames.get(k)).add(array.get(k % columnNames.size()));
			}
		}

		System.out.println(data);
	}
}
