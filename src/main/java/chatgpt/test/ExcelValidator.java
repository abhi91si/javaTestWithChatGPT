package chatgpt.test;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelValidator {

    private static final String[] COLUMN_NAMES = {"ClaimNumber", "ClaimCategory", "EmployeeCode",
        "ClaimDesc", "ReceiptDate", "ClaimedAmount", "SubmissionDate", "ClaimStatus",
        "ApprovedAmount", "ApprovalDate", "Approved By"};
    private static final int EXPECTED_NUM_COLUMNS = 11;

    public void validateExcel(String filename) throws IOException {        
        File file = new File(filename);
        
        ExcelValidator reader = new ExcelValidator();        
        reader.validateFileExtension(filename);

        // read Excel file
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);        

        reader.validateColumnNamesAndNumber(sheet);
        reader.checkDuplicateRowsAndColumns(sheet);

        //check number format
        DateAndNumberValidator.checkNumberAndDateFormat(filename);        
        workbook.close();
    }
    
    // display contents of Excel file
    public void displayFileContents(Sheet sheet) {
        for (Row row : sheet) {
            for (int i = 0; i < EXPECTED_NUM_COLUMNS; i++) {
                Cell cell = row.getCell(i);
                String cellValue = cell == null ? "" : cell.toString();
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }
    }
    
    public void checkDuplicateRowsAndColumns(Sheet sheet) {
    	Set<String> uniqueRows = new HashSet();
        for (Row row : sheet) {
            StringBuilder rowString = new StringBuilder();
            for (int i = 0; i < EXPECTED_NUM_COLUMNS; i++) {
                Cell cell = row.getCell(i);
                String cellValue = cell == null ? "" : cell.toString();
                rowString.append(cellValue);
            }
            if (uniqueRows.contains(rowString.toString())) {
                System.out.println("Error: Duplicate row found at row " + (row.getRowNum() + 1));
                return;
            }
            uniqueRows.add(rowString.toString());
        }
    }
    
    public void validateColumnNamesAndNumber(Sheet sheet) {
    	 // validate number of columns
        Row headerRow = sheet.getRow(0);
        if (headerRow == null || headerRow.getLastCellNum() != EXPECTED_NUM_COLUMNS) {
            System.out.println("Error: Invalid number of columns.");
            return;
        }

        // validate column names
        for (int i = 0; i < EXPECTED_NUM_COLUMNS; i++) {
            Cell cell = headerRow.getCell(i);
            String columnName = cell.getStringCellValue();
            if (!COLUMN_NAMES[i].equals(columnName)) {
                System.out.println("Error: Invalid column name at column " + (i + 1));
                return;
            }
        }    	
    }
    
    public void validateFileExtension(String filename) {
    	 File file = new File(filename);
        if (!file.isFile() || !filename.toLowerCase().endsWith(".xlsx")) {
            System.out.println("Error: Invalid file format or file not found.");
            System.exit(1);;
        }
    }
}

