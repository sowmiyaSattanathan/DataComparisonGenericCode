package StructureComparison.DataComparison;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFile {
	
	
	    public static XSSFSheet ConvertExcelToSheet(String excellFilePath)throws IOException {
	    	
	    	 FileInputStream fileStream = new FileInputStream(excellFilePath);
	         XSSFWorkbook workbook= new XSSFWorkbook(fileStream);
	    	XSSFSheet sheet = workbook.getSheetAt(0);
            return sheet;
		}
	    
}
