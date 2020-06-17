package StructureComparison.DataComparison;

import java.io.BufferedReader;
import java.io.FileReader;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVFile {
	
	public static XSSFSheet ConvertCSVToSheet(String csvFilePath ) {
		 XSSFSheet sheet = null;
	    try {
	        
	        XSSFWorkbook workBook = new XSSFWorkbook();
	         sheet = workBook.createSheet("sheet1");
	        String currentLine=null;
	        int RowNum=0;
	        BufferedReader br = new BufferedReader(new FileReader(csvFilePath));
	        while ((currentLine = br.readLine()) != null) {
	            String str[] = currentLine.split(",");
	           
	            XSSFRow currentRow=sheet.createRow(RowNum);
	            for(int i=0;i<str.length;i++){
	                currentRow.createCell(i).setCellValue(str[i]);
	            }
	            RowNum++;
	        }
	      
	    } catch (Exception ex) {
	        System.out.println(ex.getMessage()+"Exception in try");
	    }
	    return sheet;
	}

}
