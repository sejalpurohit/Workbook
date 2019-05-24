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

public class ReadSheetByName {
	private File file;
	private FileInputStream fileInput;
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;

	
	/**
	 * This method is used to get number of cell in the given number of row a the
	 * @param filePath- This is the first parameter to get the path of XLSX file
	 * @param sheetName- This is the second parameter to get Sheet name
	 * @param rowNum -This is the third parameter to to get the row number
	 * @throws IOException
	 */
	public void getCellNumber(String filePath, String sheetName, int rowNum) throws IOException {
		
		System.out.println("By Sheet Name-");
		long rowCount=0;
		long cellCount = 0;
		
		file = new File(filePath);
		if (file.exists() == true) { // Check id the file exist or not
			fileInput = new FileInputStream(file);

			workbook = new XSSFWorkbook(fileInput);
			int sheetNo = workbook.getSheetIndex(sheetName);
			if (workbook.getNumberOfSheets() > sheetNo ) { //check if the given sheet numeber exist in the workbook or not
														

				sheet = workbook.getSheetAt(sheetNo);
				
				Iterator<Row> rowIterator = sheet.rowIterator(); // iterate over the rows
				while (rowIterator.hasNext()) { // iterate over the cell
					Row row = rowIterator.next();
					rowCount = row.getRowNum();

					if ( row != null && row.getRowNum() == rowNum) {
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							
							Cell cell = cellIterator.next();
							cellCount = cell.getColumnIndex();
						}
					}
				}
				System.out.println("Sheet Name "+sheetName+
						"\nTotal of row"+rowCount+
						"\nNumber of Cell in Row " + (rowNum+1) + " is " + (cellCount+1));
				workbook.close();
			} else {
				System.out.println("Sheet does Not exist");
			}
		} else {
			System.out.println("File Not Found ");
	}
			
	}
}
