package chatgpt.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.System;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;


public class DateAndNumberValidator {
	
	private static final String[] COLUMN_DATE_NAMES = {"ReceiptDate", "SubmissionDate", "ApprovalDate"};

	public static void checkNumberAndDateFormat(String filename)throws EncryptedDocumentException, IOException {

        // Open the Excel file
        FileInputStream fileInputStream = new FileInputStream(new File(filename));
        Workbook workbook = WorkbookFactory.create(fileInputStream);
                
        // Get the first sheet
        Sheet sheet = workbook.getSheetAt(0);
        checkForValidAmount(fileInputStream, sheet);
        checkForValidDates(fileInputStream, sheet);
        
       // Close the Excel file
        fileInputStream.close();
        workbook.close();
    }
	
	public static void checkForValidAmount(FileInputStream fileInputStream, Sheet sheet) {
		 List<String> claim = new ArrayList<String>();
	        List<String> approve = new ArrayList<String>();
	        
	        // Iterate over the rows starting from the second row (to skip the first row)
	        Iterator<Row> rowIterator = sheet.iterator();
	        if(rowIterator.hasNext()) {
	        	rowIterator.next(); // skip first row
	        }	        	
	        
	        while (rowIterator.hasNext()) {
	            Row row = rowIterator.next();
	            
	            // Get the "ClaimedAmount" and "ApprovedAmount" cells
	            Cell claimedAmountCell = row.getCell(5); // 5 is the index of "ClaimedAmount" column
	            Cell approvedAmountCell = row.getCell(8); // 8 is the index of "ApprovedAmount" column
	            
	            // Format the cell values as strings using the DataFormatter class
	            DataFormatter dataFormatter = new DataFormatter();
	            String claimedAmountString = dataFormatter.formatCellValue(claimedAmountCell);
	            if(claimedAmountString!= null && !claimedAmountString.isBlank()) {
	            	claim.add(claimedAmountString);
	            }
	            String approvedAmountString = dataFormatter.formatCellValue(approvedAmountCell);
	            if(approvedAmountString!= null && !approvedAmountString.isBlank()) {
	            	approve.add(approvedAmountString);
	            }           
	        }
	        
	        checkForValidAmount(claim);
	        checkForValidAmount(approve);      
	}
	
	public static void checkForValidDates(FileInputStream fileInputStream, Sheet sheet) {
		for (int i = 0; i < COLUMN_DATE_NAMES.length; i++) {
			try {
				List<String> dates = ExcelUtil.getColumnValueInList(fileInputStream, sheet, COLUMN_DATE_NAMES[i], Boolean.TRUE);				
				checkForValidDates(dates);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}	
		}			
	}
	   
	//checking if List has Integer objects
    public static boolean checkForValidAmount(List<String> list) {
    	try {
    		for (String obj : list) {
    			int amount = Integer.parseInt(obj);
    			if(amount < 0) {
    				System.out.println("check the Amount value, value should be a positive number, check :"+amount);
    				System.exit(1);
    			}
    		}
    	}catch(Exception e) {
    		System.out.println("check the Amount value, value should be a number "+e.getMessage());
    	}
        return false;
    }
    
    //checking if List has strings that can be parsed as dates 
    public static void checkForValidDates(List<String> listWithDate){
    	String regex = "^(0[1-9]|[1-2][0-9]|3[01])/(0[1-9]|1[0-2])/(\\d{4})$";
    	 
    	Pattern pattern = Pattern.compile(regex);
    	 
    	for(String date : listWithDate)
    	{
    	  Matcher matcher = pattern.matcher(date);
    	  if(!matcher.matches()) {
    		 System.out.println("Date has to be in dd/mm/yyyy format, check :"+date);  
    		 System.exit(1);
    	  }
    	}
    }

}
