package excel.readExcelHashMap;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.testng.annotations.Test;

public class ExcelDataTest {

	@Test(enabled = true, priority = 1)
	public void testExcelDataUsingArrays() throws IOException {
		BasePage objBasePage = new BasePage();
		int testRowNum = objBasePage.getRowNum();
		System.out.println("Number of Rows:" + testRowNum);
		System.out.println("Method 1: Store data in array-\n");
		System.out.println(objBasePage.getExcelDataArray());
	}

	@Test(enabled = true, priority = 3)
	public void testExcelHashMapTwoMethods() throws IOException {
		BasePage objBasePage = new BasePage();
		System.out.println("Method 3: Use two methods to return the data-\n");
		System.out.println(objBasePage.getExcelDataHashMapTwoMethod());
	}

	@Test(enabled = true, priority = 2)
	public void testExcelHashMapSingleMethod() throws IOException {
		BasePage objBasePage = new BasePage();
		System.out.println("Method 2: Use single method to return the data-\n");
		System.out.println(objBasePage.getExcelDataHashMapSingleMethod());
	}

	@Test(enabled = true, priority = 4)
	public void outSideBasePage() throws IOException {
		System.out.println("Method 4: Call the method outside the base class to return the data(Error)-\n");
		BasePage objBasePage = new BasePage();
		Map<Integer, Map<String, String>> data = new HashMap<Integer, Map<String, String>>();
		for (int i = 1; i < objBasePage.getRowNum() + 1; i++) {
			if (objBasePage.emptyRow(i) == true) {
				System.out.println("Row " + i + "is a empty row.");
			} else {
				data.put(i, objBasePage.getExcelCellDataTwoMethod(i));
			}
		}
		System.out.println(data);
	}

}
