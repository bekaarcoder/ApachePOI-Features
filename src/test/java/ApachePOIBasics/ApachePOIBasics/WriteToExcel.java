package ApachePOIBasics.ApachePOIBasics;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel {

	public static void main(String[] args) throws IOException {
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employees");
		
		Object empData[][] = {
				{"EMP_ID", "NAME", "JOB"},
				{101, "John Doe", "Developer"},
				{102, "Jane Doe", "Functional Tester"},
				{103, "Cole Brown", "Automation Tester"}
		};
		
		int rows = empData.length;
		int cols = empData[0].length;
		
		for(int r=0; r<rows; r++) {
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0; c<cols; c++) {
				XSSFCell cell = row.createCell(c);
				Object cellValue = empData[r][c];
				
				if(cellValue instanceof String) {
					cell.setCellValue((String)cellValue);
				} else if(cellValue instanceof Boolean) {
					cell.setCellValue((Boolean)cellValue);
				} else if(cellValue instanceof Integer) {
					cell.setCellValue((Integer)cellValue);
				}
			}
		}
		
		FileOutputStream outputStream = new FileOutputStream(".\\dataFiles\\Employee.xlsx");
		workbook.write(outputStream);
		
		outputStream.close();
		System.out.println("Data written successfully.");
	}

}
