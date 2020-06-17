package StructureComparison.DataComparison;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import StructureComparison.DataComparison.readFile;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataCompare {
	
	public static Map < String, Object[] > compareData1 ;
    public static Map < String, Object[] > compareData2 ;
    public static Integer index = 1;
    public static XSSFSheet sheet1;
    public static XSSFSheet sheet2;
    
    
    
	 // Compare Two Sheets
    public static void compareTwoSheets(XSSFSheet _sheet1,XSSFSheet _sheet2) {  
    	compareData1 = new TreeMap < String, Object[] > ();
    	compareData2 = new TreeMap < String, Object[] > ();
    	compareData1.put(index.toString(), new Object[] {
	               "ID",
	               "Source Column Name",
	               "Source Column Value",
	               "Target Column Name",
	               "Target Column Value",
	               "Status"
	           });
	           //row missing
	           compareData2.put(index.toString(), new Object[] {
	               "ID",
	               "Status"
	           });
    	
    		sheet1= _sheet1;
    		sheet2= _sheet2;
    		
    		int firstRow1 = sheet1.getFirstRowNum();
            int lastRow1 = sheet1.getLastRowNum();

            int firstRow2 = sheet2.getFirstRowNum();
            int lastRow2 = sheet2.getLastRowNum();

            
            for (int i = firstRow1 + 1; i <= lastRow1; i++) {
                XSSFRow row1 = sheet1.getRow(i);//2
                if (row1 != null) {
                    String Row1ID = row1.getCell(0).getRawValue();
                    Boolean notFound = true;
                    for (int j = firstRow2 + 1; j <= lastRow2; j++) {
                        XSSFRow row2 = sheet2.getRow(j);
                        String Row2ID = row2.getCell(0).getRawValue();
                        if (Row1ID.equals(Row2ID)) {
                            notFound =false;
                            compareTwoRows(row1, row2, i);
                            break;
                        } 
                    }
                    if(notFound){
                        index++;
                        compareData2.put(index.toString(), new Object[] {
                           (Row1ID), "Missing row in target"
                        });
                    }
                } 
            }
            //Checking for missing rows in sheet 1
            for (int i = firstRow2 + 1; i <= lastRow2; i++) {
                XSSFRow row2 = sheet2.getRow(i);//2
                if (row2 != null) {
                    String Row2ID = row2.getCell(0).getRawValue();
                    Boolean notFound = true;
                    for (int j = firstRow1 + 1; j <= lastRow1; j++) {
                        XSSFRow row1 = sheet1.getRow(j);
                        String Row1ID = row1.getCell(0).getRawValue();
                        if (Row1ID.equals(Row2ID)) {
                            notFound =false;
                            
                            break;
                        } 
                    }
                    if(notFound){
                        index++;
                        compareData2.put(index.toString(), new Object[] {
                            (Row2ID), "Missing row in source"
                        });
                    }
                } 
            }
    	
    }

    // Compare Two Rows column
    public static void compareTwoRows(XSSFRow row1, XSSFRow row2, int rowID) {

        int firstCell1 = row1.getFirstCellNum();
        int lastCell1 = row1.getLastCellNum();

        // Compare all cells in a row
        for (int i = firstCell1; i <= lastCell1; i++) {
            XSSFCell cell1 = row1.getCell(i);
            XSSFCell cell2 = row2.getCell(i);
            if ((cell1 != null) && (cell2 != null)) {
                compareTwoCells(cell1, cell2, rowID, i);
            }
        }
    }

    // Compare Two Cells column
    public static void compareTwoCells(XSSFCell cell1, XSSFCell cell2, int rowID, int colID) {
        
        int type1 = cell1.getCellType();
        int type2 = cell2.getCellType();
        if (type1 == type2) {
            if (cell1.getCellStyle().equals(cell2.getCellStyle())) {
                // Compare cells based on its type
                switch (cell1.getCellType()) {
                    
                    case HSSFCell.CELL_TYPE_NUMERIC:
                        if (cell1.getNumericCellValue() != cell2
                            .getNumericCellValue()) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell1.getNumericCellValue(), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell2.getNumericCellValue(),"Data Mismatch"
                            });
                        } 
                            
                        
                        break;
                    case HSSFCell.CELL_TYPE_STRING:
                        if (!cell1.getStringCellValue().equals(cell2
                                .getStringCellValue())) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell1.getStringCellValue(), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell2.getStringCellValue(),"Data Mismatch"
                            });
                        } 
                        
                        break;
                    
                    case HSSFCell.CELL_TYPE_BOOLEAN:
                        if (cell1.getBooleanCellValue() != cell2
                            .getBooleanCellValue()) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getBooleanCellValue(), cell1.getBooleanCellValue(), sheet1.getRow(0).getCell(colID).getBooleanCellValue(), cell2.getBooleanCellValue(),"Data Mismatch"
                            });
                        } 
                        break;
                        
                    
                    default:
                        if (!cell1.getStringCellValue().equals(
                                cell2.getStringCellValue())) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell1.getStringCellValue(), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell2.getStringCellValue(),"Data Mismatch"
                            });
                        } 
                        break;
                }
            } 
        } 
    }
    public static void writeOutputFile(String OutputFilePath)throws IOException {
		
		   //Create blank workbook
	       XSSFWorkbook workbook = new XSSFWorkbook();	   	
		   
	           
	         
		        //Create a blank sheet
		        XSSFSheet spreadsheet1 = workbook.createSheet(" Data Mismatch ");
		        XSSFSheet spreadsheet2 = workbook.createSheet(" Data Missing ");

		        //Create row object
		        XSSFRow row;

		        //Iterate over data and write to sheet
		        Set < String > keyid1 = compareData1.keySet();
		        int rowid = 0;

		        for (String key: keyid1) {
		            row = spreadsheet1.createRow(rowid++);
		            Object[] objectArr = compareData1.get(key);
		            int cellid = 0;

		            for (Object obj: objectArr) {
		                Cell cell = row.createCell(cellid++);
		                cell.setCellValue(String.valueOf(obj));
		            }
		        }

		        //Iterate over data and write to sheet
		        Set < String > keyid2 = compareData2.keySet();
		        rowid = 0;

		        for (String key: keyid2) {
		            row = spreadsheet2.createRow(rowid++);
		            Object[] objectArr = compareData2.get(key);
		            int cellid = 0;

		            for (Object obj: objectArr) {
		                Cell cell = row.createCell(cellid++);
		                cell.setCellValue(String.valueOf(obj));
		            }
		        }
		        //Write the workbook in file system
		        FileOutputStream out = new FileOutputStream(new File(OutputFilePath));

		        workbook.write(out);

	                   
		}
    
    public static void SendMailByAttachment(String OutputFile) {
    	
    	// Recipient's email ID
        String to = Constants.To_Address;

        // Sender's email ID 
        String from = Constants.From_Address;

        final String username = Constants.User;
        final String password = Constants.Password;

        
        String host =Constants.Host;

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", "587");

        // Get the Session object.
        Session session = Session.getInstance(props,
           new javax.mail.Authenticator() {
              protected PasswordAuthentication getPasswordAuthentication() {
                 return new PasswordAuthentication(username, password);
              }
           });

        try {
           // Create a default MimeMessage object.
           Message message = new MimeMessage(session);

           // Set From: header field of the header.
           message.setFrom(new InternetAddress(from));

           // Set To: header field of the header.
           message.setRecipients(Message.RecipientType.TO,
              InternetAddress.parse(to));

           // Set Subject: header field
           message.setSubject("Automation Result");

           // Create the message part
           BodyPart messageBodyPart = new MimeBodyPart();

           // Now set the actual message
           messageBodyPart.setText("Hi,Kindly check the result output");

           // Create a multipar message
           Multipart multipart = new MimeMultipart();

           // Set text message part
           multipart.addBodyPart(messageBodyPart);

           // Part two is attachment
           messageBodyPart = new MimeBodyPart();
           String filename = OutputFile;
           DataSource source = new FileDataSource(filename);
           messageBodyPart.setDataHandler(new DataHandler(source));
           messageBodyPart.setFileName(filename);
           multipart.addBodyPart(messageBodyPart);

           // Send the complete message parts
           message.setContent(multipart);

           // Send message
           Transport.send(message);

           System.out.println("Sent message successfully....");
    
        } catch (MessagingException e) {
           throw new RuntimeException(e);
        }
     }

	
}
