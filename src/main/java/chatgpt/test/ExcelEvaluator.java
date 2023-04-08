package chatgpt.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelEvaluator {
	
	public static void evaluate(String filename) throws EncryptedDocumentException, IOException{
		 // Open the Excel file
        FileInputStream fileInputStream = new FileInputStream(new File(filename));
        Workbook workbook = WorkbookFactory.create(fileInputStream);
                
        // Get the first sheet
        Sheet sheet = workbook.getSheetAt(0);
        
        getClaimsStatus(fileInputStream, sheet);
	}

	public static void getClaimsStatus(FileInputStream fileInputStream, Sheet sheet) throws IOException {

		String columnname = "ClaimStatus";
        boolean flagForDate = Boolean.FALSE;
        
		List<String> approvalDates = ExcelUtil.getColumnValueInList(fileInputStream, sheet, columnname, flagForDate);
        System.out.println("data in list :");
        int countReview = 0;
        int countNoStatus = 0;
        int countApproved = 0;
        for (String s : approvalDates) {
        	if(s.equalsIgnoreCase("In Review")) {
        		countReview++;
        	}
        	if(s.equalsIgnoreCase("Approved")) {
        		countApproved++;
        	}
        	if(s.isEmpty()) {
        		countNoStatus++;
        	}
        }
        System.out.println("in review :"+ countReview+" no status : "+countNoStatus+ " approved :"+countApproved);

	}
}
