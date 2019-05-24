import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author sejal.purohit
 *
 */
public class ExcelRead {
	/**
	 * This method is used to get number of row in the given workbook sheet
	 * @param file     - The File object of XLSX file
	 * @param filePath -The path of XLSX file
	 * @param sheetNo
	 * @param rowNum
	 * @throws IOException
	 */
	private File file;
	private FileInputStream fileInput;
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;

	public void getLastRow(String filePath, int sheetNo) throws IOException, IllegalArgumentException {
		long rowCount = 0;

		file = new File(filePath);
		if (file.exists() == true) { // Check id the file exist or not
			fileInput = new FileInputStream(file);
			workbook = new XSSFWorkbook(fileInput);
			if (workbook.getNumberOfSheets() > sheetNo) { // check if the given sheet numebr exist in the workbook or
															// not
				sheet = workbook.getSheetAt(sheetNo);
				Iterator<Row> rowIterator = sheet.rowIterator();
				while (rowIterator.hasNext()) { // iterate over the rows
					Row row = rowIterator.next();
					rowCount = row.getRowNum();
				}
				System.out.println("Number of rows in sheet " + (sheetNo+1) + " is " + (rowCount+1));

			} else {
				System.out.println("Sheet does Not exist");
			}
		} else {
			System.out.println("File not Found");
		}
	}
	/**
	 * This method is used to get number of cell in the given number of Row
	 * 
	 * @param filePath This is the first parameter to get the path of XLSX file
	 * @param sheetNum This is the second parameter to get the workbook sheet number.
	 * @param rowNum     This is the third parameter to to get the row number
	 * @throws IOException
	 */
	public void randomCellNumber(String filePath, int sheetNo, int rowNum) throws IOException {
		long cellCount = 0;
		file = new File(filePath);
		if (file.exists() == true) { // Check id the file exist or not
			fileInput = new FileInputStream(file);

			workbook = new XSSFWorkbook(fileInput);
			if (workbook.getNumberOfSheets() > sheetNo) { //check if the given sheet numeber exist in the workbook or not
														
				sheet = workbook.getSheetAt(sheetNo);
				
				Iterator<Row> rowIterator = sheet.rowIterator(); // iterate over the rows
				while (rowIterator.hasNext()) { // iterate over the cell
					Row row = rowIterator.next();
					if ( row != null && row.getRowNum() == rowNum) {
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							
							Cell cell = cellIterator.next();
							cellCount = cell.getColumnIndex();
						}
					}
				}
				System.out.println("Number of Cell in Sheet " +(sheetNo+1)+  " Row " + (rowNum+1) + " is " + (cellCount+1));
				workbook.close();
			} else {
				System.out.println("Sheet does Not exist");
			}
		} else {
			System.out.println("File Not Found ");
	}

	}
}
