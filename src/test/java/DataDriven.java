import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class DataDriven {

	private static final String absolutePath = "/Users/kmieshkov/Projects/IdeaProjects/maven-excel";

	public HashMap<String, String> getFullData(String fileName, String sheetName, String testCaseId) throws IOException {
		FileInputStream fileInputStream = new FileInputStream(absolutePath + fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

		ArrayList<String> columnNames = new ArrayList<>();
		HashMap<String, String> data = new HashMap<>();

		XSSFSheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
		}
		Iterator<Row> rows = sheet.iterator();

		// iterator for each row in a sheet
		for (int j = 0; rows.hasNext(); j++) {
			Row row = rows.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			ArrayList<String> array = new ArrayList<>();

			// iterator for each cell in a row\
			int needleRow = 0;
			while (cellIterator.hasNext()) {
				Cell cellValue = cellIterator.next();
				if (j == 0) {
					columnNames.add(formattedCellValue(cellValue));
				} else if (formattedCellValue(cellValue).equals(testCaseId) || needleRow == 1) {
					needleRow = 1;
					array.add(formattedCellValue(cellValue));
				}
			}

			// save data in format "["ID": "TC-01", "Test Data": "Data123", ...]
			for (int k = 0; k < array.size(); k++) {
				System.out.printf("%s = %s\n", columnNames.get(k), array.get(k));
				data.put(columnNames.get(k), array.get(k));
			}
		}
		return(data);
	}

	public ArrayList<String> getData(String fileName, String sheetName, String rowName) throws IOException {
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fileInputStream = new FileInputStream(absolutePath + fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		ArrayList<String> array = new ArrayList<>();

		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				//Identify Testcases column by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();// sheet is collection of rows
				Row firstRow = rows.next();
				Iterator<Cell> cellIterator = firstRow.cellIterator();//row is collection of cells

				// Find the column with names
				int column = 0;
				for (int k = 0; cellIterator.hasNext(); k++) {
					Cell value = cellIterator.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column = k;
					}
				}
				while (rows.hasNext()) {
					Row row = rows.next();

					// grab all cells for Purchase
					if (row.getCell(column).getStringCellValue().equalsIgnoreCase(rowName)) {
						Iterator<Cell> cell = row.cellIterator();
						while (cell.hasNext()) {
							array.add(formattedCellValue(cell.next()));
						}
					}
				}
			}
		}
		return array;
	}

	private String formattedCellValue(Cell cell) {
		if (cell.getCellType() == CellType.STRING) {
			return cell.getStringCellValue();
		} else {
			return NumberToTextConverter.toText(cell.getNumericCellValue());
		}
	}

}
