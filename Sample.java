package org.sag;
	import java.io.File;
import java.io.IOException;
	import java.math.BigDecimal;
	import java.text.SimpleDateFormat;
	import java.util.Date;

	import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.CellType;
	import org.apache.poi.ss.usermodel.DateUtil;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


	public class Sample {
		public static void main(String[] args) throws InvalidFormatException, IOException {
			File file = new File("C:\\Users\\AswinRock\\eclipse-workspace123\\Data1\\Xl\\Book1.xlsx");
			Workbook workbook = new XSSFWorkbook(file);
			Sheet sheet = workbook.getSheet("sheet1");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j <row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				
				CellType type = cell.getCellType();
				switch (type) {
				case STRING:
					String cellValue = cell.getStringCellValue();
					System.out.println(cellValue);
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						Date value = cell.getDateCellValue();
						SimpleDateFormat format=new SimpleDateFormat("DD/MM/YYYY");
						String format2 = format.format(value);
						System.out.println(format2);
					
					}
					else {
						
						double cellValue2 = cell.getNumericCellValue();
						BigDecimal decimal = new BigDecimal(cellValue2 );
						String string = decimal.toString();
						System.out.println(string);	
					}
				break;
				}
				
				
			}
			System.out.println("");
		}
		
			
		}

	}



