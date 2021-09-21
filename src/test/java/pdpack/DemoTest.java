package pdpack;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DemoTest {

	public static void main(String[] args) throws IOException {
		ArrayList<String> alist = getDataFromExcelFile("Login" ,"C:\\Users\\nbhar\\OneDrive\\Desktop\\ExcelTestData.xlsx","SheetA");
		for(String a:alist) {
			System.out.println(a);
		}
	}

	public static ArrayList<String> getDataFromExcelFile(String testName ,String pathOfexcelFile,String sheetName) throws IOException {

		ArrayList<String> al = new ArrayList<String>();
		FileInputStream fis = new FileInputStream(pathOfexcelFile);
		XSSFWorkbook wrkbook = new XSSFWorkbook(fis);
		int sheetcount = wrkbook.getNumberOfSheets();

		for (int i = 0; i < sheetcount; i++) {
			if (wrkbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
				XSSFSheet sheet = wrkbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();
				Iterator<Cell> firstrowCells = firstrow.iterator();

				int c = 0;
				int TestColumnPosition = 0;
				while (firstrowCells.hasNext()) {
					// System.out.println(firstrowCells.next().getStringCellValue());
					Cell firstRowCell = firstrowCells.next();
					//System.out.println(firstRowCell.getStringCellValue());
					if (firstRowCell.getStringCellValue().equalsIgnoreCase("Tests")) {
						//System.out.println(firstRowCell.getStringCellValue());
						TestColumnPosition = c;
					}
					c++;
				}
				while (rows.hasNext()) {
					Row row = rows.next();
					Cell cell = row.getCell(TestColumnPosition);
					if (cell.getStringCellValue().equalsIgnoreCase(testName)) {
						Iterator<Cell> cells = row.iterator();
						cells.next();
						while (cells.hasNext()) {

							Cell currentcell = cells.next();
							if (currentcell.getCellType() == CellType.STRING) {
								// System.out.println(currentcell.getStringCellValue());
								al.add(currentcell.getStringCellValue());
							} else if (currentcell.getCellType() == CellType.NUMERIC) {
								// System.out.println(currentcell.getNumericCellValue());
								al.add(NumberToTextConverter.toText(currentcell.getNumericCellValue()));
							}
						}
					}
				}
			}

		}
		return al;
		
	}

}
