package StructureComparison.DataComparison;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.testng.annotations.Test;
import StructureComparison.DataComparison.readFile;
import StructureComparison.DataComparison.DataCompare;
import StructureComparison.DataComparison.Database;

public class MainFile {
	 @Test
	 public void readsoucefromdb() throws IOException{
		 Database exporter = new Database();
	        XSSFSheet sheet1 = exporter.ExportDbTableToSheet("customerdeatails","select * from customerdeatails");
	        XSSFSheet sheet2 = exporter.ExportDbTableToSheet("orderdetails","select * from orderdetails");
	        DataCompare.compareTwoSheets(sheet1,sheet2);
			DataCompare.writeOutputFile("E:\\my workspace\\testdb\\DataComparison\\Excel\\DBOutput.xlsx");
	 }
	
	@Test
	public void readsoucefromexcel() throws IOException {
		XSSFSheet sheet1 = readFile.ConvertExcelToSheet("E:\\my workspace\\testdb\\DataComparison\\Excel\\SourceFile.xlsx");
		XSSFSheet sheet2 = readFile.ConvertExcelToSheet("E:\\my workspace\\testdb\\DataComparison\\Excel\\TragetFile.xlsx");
		DataCompare.compareTwoSheets(sheet1,sheet2);
		DataCompare.writeOutputFile("E:\\my workspace\\testdb\\DataComparison\\Excel\\ExcelOutput.xlsx");
	}
	
	@Test
	public void readsoucefromcsv() throws IOException {
		XSSFSheet sheet1 = CSVFile.ConvertCSVToSheet("E:\\my workspace\\testdb\\DataComparison\\Excel\\SourceFileCSV.csv");
		XSSFSheet sheet2 = CSVFile.ConvertCSVToSheet("E:\\my workspace\\testdb\\DataComparison\\Excel\\TragetFileCSV.csv");
		DataCompare.compareTwoSheets(sheet1,sheet2);
		DataCompare.writeOutputFile("E:\\my workspace\\testdb\\DataComparison\\Excel\\CSVOutput.xlsx");
	}
	

}
