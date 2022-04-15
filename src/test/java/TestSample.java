import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class TestSample {

	public static void main(String[] adrgs) throws IOException {
		DataDriven dataDriven = new DataDriven();

		// get data in format Key => value for the whole row
		HashMap<String, String> data1 = dataDriven.getFullData("/testcases.xlsx", "TestCases", "TC-03");
		HashMap<String, String> data2 = dataDriven.getFullData("/testdata.xlsx", "TestData", "Purchase");

		//get data in array format [rowName, data0, data1, data2] for the row (with row name)
		ArrayList<String> data3 = dataDriven.getData("/testdata.xlsx", "TestData", "Add Profile");
		ArrayList<String> data4 = dataDriven.getData("/testcases.xlsx", "TestCases", "TC-03");
		System.out.println(data1);
		System.out.println(data2);
		System.out.println(data3);
		System.out.println(data4);
	}
}
