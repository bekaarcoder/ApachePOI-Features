package ApachePOIBasics.ApachePOIBasics;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFormulaCell {

	public static void main(String[] args) throws IOException {
		
		FileInputStream inputStream = new FileInputStream(".\\dataFiles\\Formula.xlsx");
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		XSSFSheet sheet = workbook.getSheet("Salary");
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(0).getLastCellNum();
		
		for(int r=0; r<=rows; r++) {
			XSSFRow row = sheet.getRow(r);
			
			for(int c=0; c<cols; c++) {
				XSSFCell cell = row.getCell(c);
				DataFormatter formatter = new DataFormatter();
				
				switch(cell.getCellType()) {
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: 
					System.out.print(formatter.formatCellValue(cell));
					break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				case FORMULA: 
					System.out.print(cell.getNumericCellValue());
					break;
				default: System.out.print(cell.getRawValue());
					break;				
				}
				
				System.out.print(" | ");
			}
			System.out.println();
		}
		inputStream.close();

	}

}
