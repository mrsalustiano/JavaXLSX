package br.com.marcelo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LerXLSX {

	public static void main(String[] args) {
		
		File file = new File("C:/Users/m_sal/Desktop/java/LerExcel/planilhaDaAula.xlsx");
		try {
			FileInputStream fisPlanilha = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fisPlanilha);
			
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			Iterator<Row> rowIterator = sheet.iterator();
			
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				
				Iterator<Cell> cellIterator = row.iterator();
				
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					
					switch(cell.getCellType()) {
					
					case Cell.CELL_TYPE_STRING:
						System.out.println("Tipo String: " + cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println("Tipo Numerico: " + cell.getNumericCellValue());
						break;
						
					case Cell.CELL_TYPE_FORMULA:
						System.out.println("Tipo Formula: " + cell.getCellFormula());
						break;
						
					}
				}
				
			}
					
		
		
		
		
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		

	}
}	
		
		
		