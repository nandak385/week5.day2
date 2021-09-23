package week5.day2;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel1 {

	public static String[][] readData(String sheetName) throws IOException {
		// Step1 : Locate the Workbook(setup the path)
		XSSFWorkbook wb = new XSSFWorkbook("./data/Leads.xlsx");

		// Step2 : Get into the sheet
		XSSFSheet ws = wb.getSheet(sheetName);

		int rowCount = ws.getLastRowNum();

		short cellCount = ws.getRow(0).getLastCellNum();

		String data[][] = new String[rowCount][cellCount];
		
		for (int i = 1; i <= rowCount; i++) {
			for (int j = 0; j < cellCount; j++) {

				String text = ws.getRow(i).getCell(j).getStringCellValue();
				System.out.println(text);
				data[i-1][j] = text;
				
			}
			

		}
		

		// last step
		wb.close();
		return data;
	}

}