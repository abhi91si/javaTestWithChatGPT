package chatgpt.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DateAndNumberValidator {

	private static final String[] COLUMN_DATE_NAMES = { "ReceiptDate", "SubmissionDate", "ApprovalDate" };

	public void checkNumberAndDateFormat(String filename)
			throws EncryptedDocumentException, IOException, ExcelException {

		FileInputStream fileInputStream = null;
		Workbook workbook = null;
		try {
			// Open the Excel file
			fileInputStream = new FileInputStream(new File(filename));
			workbook = WorkbookFactory.create(fileInputStream);

			// Get the first sheet
			Sheet sheet = workbook.getSheetAt(0);
			checkForValidAmount(fileInputStream, sheet);
			checkForValidDates(fileInputStream, sheet);
		} finally {
			// Close the Excel file
			if(fileInputStream != null) fileInputStream.close();
			if(workbook != null) workbook.close();
		}
	}

	private void checkForValidAmount(FileInputStream fileInputStream, Sheet sheet) throws ExcelException {
		List<String> claim = new ArrayList<String>();
		List<String> approve = new ArrayList<String>();

		// Iterate over the rows starting from the second row (to skip the first row)
		Iterator<Row> rowIterator = sheet.iterator();
		if (rowIterator.hasNext()) {
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
			if (claimedAmountString != null && !claimedAmountString.isBlank()) {
				claim.add(claimedAmountString);
			}
			String approvedAmountString = dataFormatter.formatCellValue(approvedAmountCell);
			if (approvedAmountString != null && !approvedAmountString.isBlank()) {
				approve.add(approvedAmountString);
			}
		}
		checkForValidAmount(claim);
		checkForValidAmount(approve);
	}

	private void checkForValidDates(FileInputStream fileInputStream, Sheet sheet) throws ExcelException {
		for (int i = 0; i < COLUMN_DATE_NAMES.length; i++) {
			try {
				List<String> dates = ExcelUtil.getColumnValueInList(fileInputStream, sheet, COLUMN_DATE_NAMES[i],
						Boolean.TRUE);
				checkForValidDates(dates);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	// checking if List has Integer objects
	private void checkForValidAmount(List<String> list) throws ExcelException {
		for (String obj : list) {
			try {
				int amount = Integer.parseInt(obj);
				if (amount < 0) {
					throw new ExcelException(
							"check the Amount value, value should be a positive number, check :" + amount);
				}
			} catch (NumberFormatException e) {
				throw new ExcelException("check the Amount value, " + e.getMessage());
			}
		}

	}

	// checking if List has strings that can be parsed as dates
	private void checkForValidDates(List<String> listWithDate) throws ExcelException {
		String regex = "^(0[1-9]|[1-2][0-9]|3[01])/(0[1-9]|1[0-2])/(\\d{4})$";

		Pattern pattern = Pattern.compile(regex);

		for (String date : listWithDate) {
			Matcher matcher = pattern.matcher(date);
			if (!matcher.matches()) {
				throw new ExcelException("Date has to be in dd/mm/yyyy format, Also check date, month and year, CHECK :" + date);
			}
		}
	}

}
