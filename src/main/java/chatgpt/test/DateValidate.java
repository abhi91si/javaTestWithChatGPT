package chatgpt.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
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


public class DateValidate {

	public static void checkNumberAndDateFormat(String filename)throws EncryptedDocumentException, IOException {

        // Open the Excel file
        FileInputStream fileInputStream = new FileInputStream(new File(filename));
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        List<String> claim = new ArrayList<String>();
        List<String> approve = new ArrayList<String>();
        List<String> receiptDate = new ArrayList<String>();
        
        // Get the first sheet
        Sheet sheet = workbook.getSheetAt(0);
        
        // Iterate over the rows starting from the second row (to skip the first row)
        Iterator<Row> rowIterator = sheet.iterator();
        if(rowIterator.hasNext()) rowIterator.next(); // skip first row
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            // Get the "ClaimedAmount" and "ApprovedAmount" cells
            Cell claimedAmountCell = row.getCell(5); // 5 is the index of "ClaimedAmount" column
            Cell approvedAmountCell = row.getCell(8); // 8 is the index of "ApprovedAmount" column
            Cell receiptDateCell = row.getCell(4);
            
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
            String receiptDateString = dataFormatter.formatCellValue(receiptDateCell);
            if(receiptDateString!= null && !receiptDateString.isBlank()) {
            	receiptDate.add(receiptDateString);
            }
        }
        System.out.println(claim);
        System.out.println(approve);
        System.out.println(receiptDate);
        
        isListDouble(claim);
        isListDouble(approve);
        isListDate(receiptDate);        
        
        
       // Close the Excel file
        fileInputStream.close();
        workbook.close();
    }
    
    public static boolean isListDouble(List<String> list) {
    	try {
    		for (String obj : list) {
    			Integer.parseInt(obj);
    		}
    	}catch(Exception e) {
    		System.out.println("check amount value "+e.getMessage());
    	}
        return false;
    }
    
    public static boolean isListDate(List<?> list) {
        DateFormat dateFormat1 = new SimpleDateFormat("dd/MM/yyyy");
        DateFormat dateFormat2 = new SimpleDateFormat("dd-MM-yyyy");
        for (Object obj : list) {
            if (!(obj instanceof String)) {
                return false;
            }
            String dateString = (String) obj;
            try {
                // Try parsing the string as a date using the two date formats
                dateFormat1.parse(dateString);
            } catch (ParseException e) {
                try {
                    dateFormat2.parse(dateString);
                } catch (ParseException e2) {
                    // If parsing fails, the string is not a valid date
                	System.out.println("value : "+ dateString);
                    return false;
                }
            }
        }
        // If all strings are valid dates, return true
        return true;
    }
}
