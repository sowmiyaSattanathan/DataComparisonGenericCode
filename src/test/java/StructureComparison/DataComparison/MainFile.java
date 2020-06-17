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
	 public void readsourcefromdb() throws IOException{
		 Database exporter = new Database();
	        XSSFSheet sheet1 = exporter.ExportDbTableToSheet("customerdeatails","select * from customerdeatails");
	        XSSFSheet sheet2 = exporter.ExportDbTableToSheet("orderdetails","select * from orderdetails");
	        DataCompare.compareTwoSheets(sheet1,sheet2);
			DataCompare.writeOutputFile("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\DBOutput.xlsx");
			DataCompare.SendMailByAttachment("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\DBOutput.xlsx");
	 }
	
	@Test
	public void readsourcefromexcel() throws IOException {
		XSSFSheet sheet1 = readFile.ConvertExcelToSheet("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\SourceFile.xlsx");
		XSSFSheet sheet2 = readFile.ConvertExcelToSheet("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\TragetFile.xlsx");
		DataCompare.compareTwoSheets(sheet1,sheet2);
		DataCompare.writeOutputFile("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\ExcelOutput.xlsx");
		DataCompare.SendMailByAttachment("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\ExcelOutput.xlsx");
	}
	
	@Test
	public void readsoucefromcsv() throws IOException {
		XSSFSheet sheet1 = CSVFile.ConvertCSVToSheet("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\SourceFileCSV.csv");
		XSSFSheet sheet2 = CSVFile.ConvertCSVToSheet("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\TragetFileCSV.csv");
		DataCompare.compareTwoSheets(sheet1,sheet2);
		DataCompare.writeOutputFile("F:\\DataComparison_Update_Testng\\DataComparisonGenericCode\\Excel\\CSVOutput.xlsx");
	}
	
	

}
