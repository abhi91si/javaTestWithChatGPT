package chatgpt.test;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelValidator {

	private static final String[] COLUMN_NAMES = { "ClaimNumber", "ClaimCategory", "EmployeeCode", "ClaimDesc",
			"ReceiptDate", "ClaimedAmount", "SubmissionDate", "ClaimStatus", "ApprovedAmount", "ApprovalDate",
			"Approved By" };
	private static final int EXPECTED_NUM_COLUMNS = 11;

	public void validateExcel(String filename) throws IOException, ExcelException {
		Workbook workbook = null;
		Sheet sheet = null;
		try {
			ExcelValidator reader = new ExcelValidator();
			reader.validateFileExtension(filename);
			System.out.println("File extension is validated");
			File file = new File(filename);

			// read Excel file
			try {
				workbook = WorkbookFactory.create(file);
				sheet = workbook.getSheetAt(0);
				if (sheet.getLastRowNum() <= 0) {
					throw new ExcelException("File is empty");
				}
			} catch (Exception e) {
				throw new ExcelException("Please close the file if openned or Check if file is a valid Excel file !!!");
			} finally {
				if (workbook != null)
					workbook.close();
			}

			reader.validateColumnNames(sheet);
			System.out.println("Column names are validated");
			reader.validateColumnNumber(sheet);
			System.out.println("Column numbers are validated");
			reader.checkDuplicateRowsAndColumns(sheet);
			System.out.println("If duplicate data exists is validated");

			// check number format
			DateAndNumberValidator dateAndNumberValidator = new DateAndNumberValidator();
			dateAndNumberValidator.checkNumberAndDateFormat(filename);
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} finally {
			if (workbook != null)
				workbook.close();
		}
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

	public void checkDuplicateRowsAndColumns(Sheet sheet) throws ExcelException {
		Set<String> uniqueRows = new HashSet();
		for (Row row : sheet) {
			StringBuilder rowString = new StringBuilder();
			for (int i = 0; i < EXPECTED_NUM_COLUMNS; i++) {
				Cell cell = row.getCell(i);
				String cellValue = cell == null ? "" : cell.toString();
				rowString.append(cellValue);
			}
			if (uniqueRows.contains(rowString.toString())) {
				throw new ExcelException("Error: Duplicate row found at row " + (row.getRowNum() + 1));
			}
			uniqueRows.add(rowString.toString());
		}
	}

	public void validateColumnNames(Sheet sheet) throws ExcelException {
		// validate number of columns
		Row headerRow = sheet.getRow(0);
		if (headerRow == null || headerRow.getLastCellNum() != EXPECTED_NUM_COLUMNS) {
			throw new ExcelException("Error: Invalid number of columns.");
		}
	}

	public void validateColumnNumber(Sheet sheet) throws ExcelException {
		Row headerRow = sheet.getRow(0);
		// validate column names
		for (int i = 0; i < EXPECTED_NUM_COLUMNS; i++) {
			Cell cell = headerRow.getCell(i);
			String columnName = cell.getStringCellValue();
			if (!COLUMN_NAMES[i].equals(columnName)) {
				throw new ExcelException("Error: Invalid column name at column " + (i + 1));
			}
		}
	}

	public void validateFileExtension(String filename) throws ExcelException {
		File file = new File(filename);
		if (!file.isFile() || !filename.toLowerCase().endsWith(".xlsx")) {
			throw new ExcelException("Error: Invalid file format or file not found.");
		}
	}
}
