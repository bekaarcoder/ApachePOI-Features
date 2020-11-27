package ApachePOIBasics.ApachePOIBasics;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream inputStream = new FileInputStream(".\\dataFiles\\Population.xlsx");
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("World Population");
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(0).getLastCellNum();
		
		for(int r=0; r<=rows; r++) {
			XSSFRow row = sheet.getRow(r);
			
			for(int c=0; c<cols; c++) {
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType()) {
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: 
					DataFormatter formatter = new DataFormatter();
					String cellValue = formatter.formatCellValue(cell);
					System.out.print(cellValue);
					break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
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
