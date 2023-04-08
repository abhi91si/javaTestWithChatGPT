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

public class ExcelReader {

    private static final String[] COLUMN_NAMES = {"ClaimNumber", "ClaimCategory", "EmployeeCode",
        "ClaimDesc", "ReceiptDate", "ClaimedAmount", "SubmissionDate", "ClaimStatus",
        "ApprovedAmount", "ApprovalDate", "Approved By"};
    private static final int EXPECTED_NUM_COLUMNS = 11;

    public static void main(String[] args) throws IOException {
        String filename = "C:\\Users\\hp\\Documents\\test.xlsx"; // replace with your file name
        File file = new File(filename);

        // validate file format
        if (!file.isFile() || !filename.toLowerCase().endsWith(".xlsx")) {
            System.out.println("Error: Invalid file format or file not found.");
            return;
        }

        // read Excel file
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

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

        // check for duplicate rows and columns
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
        //check number format
        DateAndNumberValidator.checkNumberAndDateFormat(filename);

        // display contents of Excel file
        for (Row row : sheet) {
            for (int i = 0; i < EXPECTED_NUM_COLUMNS; i++) {
                Cell cell = row.getCell(i);
                String cellValue = cell == null ? "" : cell.toString();
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        workbook.close();
    }
}

