import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.InputMismatchException;
/**
 * 
 * @author sejal.purohit This class is used to call the service from ExcelRead Class.
 *        
 *
 */
public class MainClass {

	/**
	 * 
	 * @param filePath -The path of XLSX file
	 * @param sheetNo - To get the Sheet Number
	 * @param rowNum - To get The Row Number
	 * @param sheetname - to get the Sheet Name
	 * @throws IOException If an input or output exception occurred
	 * @throwsInputMismatchException when argument doesn’t match the expected pattern or type.
	 * @throwIllegalArgumentException to indicate that a method has an illegal or  inappropriate argument.
	 *                               
	 */
	static private String filePath;
	static private int sheetNum;
	static private int rowNum;
	static private String sheetName;
	static private BufferedReader reader;

	public static void main(String[] args) {
		ExcelRead read = new ExcelRead();
		ReadSheetByName readSheet = new ReadSheetByName();	
		try {
			reader= new BufferedReader(new InputStreamReader(System.in));
			
			System.out.println("Enter File Path");
			filePath = reader.readLine();   //Reading XLSX File path
			
			System.out.println("Enter Sheet Number");
			sheetNum = Integer.parseInt(reader.readLine()); //Reading Sheet Number
			
			System.out.println("Enter Row Number");
			rowNum = Integer.parseInt(reader.readLine());  //Reading Row Number
			
			System.out.println("Enter Sheet Name");  //Reading Sheet Name
			sheetName = reader.readLine();
			
			read.getLastRow(filePath, (sheetNum-1));    //invoking getLastRow Method of ExcelRead Class
			
			read.randomCellNumber(filePath, (sheetNum-1), (rowNum-1));  //invoking randomCellNumber Method of ExcelRead Class
			
			readSheet.getCellNumber(filePath, sheetName, rowNum-1); //invoking get getCellNumber Method of ReadSheetByName
			
		} catch (IOException e) {
			System.out.println("Input Mismatch");
		} catch (InputMismatchException e) {
			System.out.println("InValid Input");
		} catch (IllegalArgumentException e) {
			System.out.println("Invalid Number");
		}		
	}
}