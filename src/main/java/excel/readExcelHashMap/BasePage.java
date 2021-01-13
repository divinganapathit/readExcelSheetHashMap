package excel.readExcelHashMap;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BasePage {
	public FileInputStream fileRead;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public XSSFRow row;
	public XSSFCell cell;
	int rowNum;
	int cellNum;
	String[] cellValue;
	Map<String, String> dataEachRow = new HashMap<String, String>();
	Map<Integer, Map<String, String>> excelData = new HashMap<Integer, Map<String, String>>();

	/*
	 * 
	 * Constructor to initialize file reading and workspace
	 * 
	 */
	public BasePage() throws IOException {
		fileRead = new FileInputStream(
				new File(System.getProperty("user.dir") + "\\src\\main\\resource\\testData.xlsx"));
		workbook = new XSSFWorkbook(fileRead);
		sheet = workbook.getSheetAt(0);
	}

	/*
	 * 
	 * Method to check whether the row is empty
	 * 
	 */
	public boolean emptyRow(int rowNum) {
		Row rowCheck = sheet.getRow(rowNum);
		for (int c = rowCheck.getFirstCellNum(); c < rowCheck.getLastCellNum(); c++) {
			Cell cellCheck = rowCheck.getCell(c);
			if (cellCheck != null && cellCheck.getCellType() != CellType.BLANK) {
				return false;
			}
		}
		return true;
	}

	/*
	 * 
	 * Method to get the number of rows
	 * 
	 */
	public int getRowNum() {
		rowNum = sheet.getLastRowNum();
		return rowNum;
	}

	/*
	 * 
	 * 
	 * Method to get the number of columns/cells
	 * 
	 */
	public int getCellNum() {
		row = sheet.getRow(0);
		cellNum = row.getLastCellNum();
		return cellNum;
	}

	/*
	 * 
	 * 
	 * Method to store the data in the array
	 * 
	 */
	public String[] getValue(int rowValue) {
		row = sheet.getRow(rowValue);
		cellNum = row.getLastCellNum();
//		System.out.println("Number of Cell:"+cellNum);
		cellValue = new String[cellNum];
		for (int i = 0; i < cellNum; i++) {
			Cell cell = row.getCell(i);
			switch (cell.getCellType()) {
			case STRING:
				cellValue[i] = String.valueOf(cell.getStringCellValue());
				break;
			case NUMERIC:
				cellValue[i] = String.valueOf(cell.getNumericCellValue());
				break;
			case FORMULA:
				cellValue[i] = String.valueOf(cell.getCellFormula());
				break;
			case BLANK:
				cellValue[i] = " ";
				break;
			default:
				System.out.println(cell.getStringCellValue());
			}
		}
		return cellValue;
	}

	/*
	 * 
	 * Method to wherein the data in the array from the above method is put into the
	 * hashmap
	 * 
	 * 
	 */
	public Map<Integer, Map<String, String>> getExcelDataArray() throws IOException {
		try {
			BasePage objBasePage = new BasePage();
			System.out.println(objBasePage.getRowNum());
			String[] header = new String[objBasePage.getCellNum()];
			String[] cellData = new String[objBasePage.getCellNum()];
			for (int i = 0; i < objBasePage.getRowNum(); i++) {
				if (objBasePage.emptyRow(i) == true) {
					System.out.println("Row " + i + "is a empty row.");
				} else {
					header = objBasePage.getValue(0);
					cellData = objBasePage.getValue(i + 1);
					for (int j = 0; j < objBasePage.getCellNum(); j++) {
						dataEachRow.put(header[j], cellData[j]);
					}
					excelData.put(i + 1, new HashMap<>(dataEachRow));
				}
			}
		} catch (Exception e) {
			System.out.println(e);
		} finally {
			fileRead.close();
			workbook.close();
		}

		return excelData;
	}

	/*
	 * 
	 * 
	 * Method wherein the cell data is store into hashmap and later put into another
	 * hashmap which stores the row data
	 * 
	 */
	public Map<Integer, Map<String, String>> getExcelDataHashMapSingleMethod() throws IOException {
		try {
			BasePage objBasePage = new BasePage();
			for (int i = 1; i < objBasePage.getRowNum() + 1; i++) {
				if (objBasePage.emptyRow(i) == true) {
					System.out.println("Row " + i + "is a empty row.");
				} else {
					for (int j = 0; j < objBasePage.getCellNum(); j++) {
						Cell cell = sheet.getRow(i).getCell(j);
						switch (cell.getCellType()) {
						case STRING:
							dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()),
									String.valueOf(cell.getStringCellValue()));
							System.out.println("String:" + dataEachRow);
							break;
						case NUMERIC:
							dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()),
									String.valueOf(cell.getNumericCellValue()));
							System.out.println("Numeric:" + dataEachRow);
							break;
						case FORMULA:
							dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()),
									String.valueOf(cell.getCellFormula()));
							System.out.println("Formula:" + dataEachRow);
							break;
						case BLANK:
							System.out.println(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()));
							dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()), " ");
							System.out.println("Blank" + dataEachRow);
							break;
						default:
							System.out.println(cell.getStringCellValue());
						}
					}
					excelData.put(i, new HashMap<>(dataEachRow));
				}
			}
		} catch (Exception e) {
			System.out.println(e);
		} finally {
			fileRead.close();
			workbook.close();
		}
		return excelData;
	}

	/*
	 * 
	 * 
	 * Method which returns the data in the cell in a hashmap
	 * 
	 */
	public Map<String, String> getExcelCellDataTwoMethod(int rowNumber) throws IOException {
		BasePage objBasePage = new BasePage();
		for (int j = 0; j < objBasePage.getCellNum(); j++) {
			Cell cell = sheet.getRow(rowNumber).getCell(j);
			switch (cell.getCellType()) {
			case STRING:
				dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()),
						String.valueOf(cell.getStringCellValue()));
				break;
			case NUMERIC:
				dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()),
						String.valueOf(cell.getNumericCellValue()));
				break;
			case FORMULA:
				dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()),
						String.valueOf(cell.getCellFormula()));
				break;
			case BLANK:
				System.out.println(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()));
				dataEachRow.put(String.valueOf(sheet.getRow(0).getCell(j).getStringCellValue()), " ");
				break;
			default:
				System.out.println(cell.getStringCellValue());
			}
		}
		return dataEachRow;
	}

	/*
	 * 
	 * Method that adds data from the hashmap returned from the above method, and
	 * returns the hashmap that stores the data in the row
	 * 
	 * 
	 */
	public Map<Integer, Map<String, String>> getExcelDataHashMapTwoMethod() throws IOException {
		try {
			BasePage objBasePage = new BasePage();
			for (int i = 1; i < objBasePage.getRowNum() + 1; i++) {
				if (objBasePage.emptyRow(i) == true) {
					System.out.println("Row " + i + "is empty row.");
				} else {
					Map<String, String> temp = new HashMap<String, String>();
					temp = objBasePage.getExcelCellDataTwoMethod(i);
					excelData.put(i, new HashMap<>(temp));
				}
			}
		} catch (Exception e) {
			System.out.println(e);
		} finally {
			fileRead.close();
			workbook.close();
		}
		return excelData;
	}
}
