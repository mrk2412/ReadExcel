package ReadExclFile;

import java.io.FileInputStream; 
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class ReadExclFile {

	public static void main(String[] args) throws IOException {
    String excelFilePath=".\\Book1.xlsx";
    FileInputStream inputStream=new FileInputStream(excelFilePath);
    
    XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
    XSSFSheet sheet=workbook.getSheetAt(0);
    
    Iterator iterator=sheet.iterator();
    
    while(iterator.hasNext())
    {
        XSSFRow	row=(XSSFRow) iterator.next();
    
        Iterator cellIterator=row.cellIterator();
        while(cellIterator.hasNext())
        {
        	XSSFCell cell=(XSSFCell) cellIterator.next();
        	
        	 switch(cell.getCellType())
        	 {
        	 case STRING: System.out.print(cell.getStringCellValue()); break;
        	 case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
        	 case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
        	 }
        	 System.out.print("  |  ");
        }
        System.out.println();
      }  
	}
}
