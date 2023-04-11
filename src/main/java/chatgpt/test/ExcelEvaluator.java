package chatgpt.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import java.text.DecimalFormat;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;

import java.util.Map;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

public class ExcelEvaluator {

	public static void evaluate(String filename) throws EncryptedDocumentException, IOException {

		// Open the Excel file
		FileInputStream fileInputStream = new FileInputStream(new File(filename));
		Workbook workbook = WorkbookFactory.create(fileInputStream);

		// Get the first sheet
		Sheet sheet = workbook.getSheetAt(0);

		Scanner input = new Scanner(System.in);

		do {
			System.out.println("\n Press '1' to get Claims Status(Approved, In review and No status) "
					+ "\n Press '2' to get Category wise total approved claim amount. "
					+ "\n Press '3' to get Monthly/Quarterly/Yearly claims submission vs approvals."
					+ "\n Press '4' to get Projected Category wise claims for the next quarter, based on the current trend."
					+ "\n Press '5' to get all the evaluations.");

			System.out.print("Enter your choice : ");
			int choice = input.nextInt();

			switch (choice) {
			case 1:
				getClaimsStatus(fileInputStream, sheet);
				break;
			case 2:
				getCategoryWiseClaims(filename);
				break;
			case 3:
				getClaimApprovedAndSubmittedStatus(filename);
				break;
			case 4:
				claimProjections(filename);
				break;
			case 5:
				getClaimsStatus(fileInputStream, sheet);
				getCategoryWiseClaims(filename);
				getClaimApprovedAndSubmittedStatus(filename);
				claimProjections(filename);
				break;
			default:
				System.out.println("Invalid input. Please enter a number between 1 and 5.");
				break;
			}
			System.out.print("\nDo you want to see more evaluations ? (y/n): ");
		} while (input.next().equalsIgnoreCase("y"));
		System.out.println("Thank you !!");
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
			if (s.equalsIgnoreCase("In Review")) {
				countReview++;
			}
			if (s.equalsIgnoreCase("Approved")) {
				countApproved++;
			}
			if (s.isEmpty()) {
				countNoStatus++;
			}
		}
		System.out.println(
				"in review :" + countReview + " no status : " + countNoStatus + " approved :" + countApproved + "\n");
	}

	public static void getClaimApprovedAndSubmittedStatus(String file) throws IOException {

		FileInputStream inputStream = new FileInputStream(file);
		Workbook workbook = WorkbookFactory.create(inputStream);

		// Get sheet and rows
		Sheet sheet = workbook.getSheetAt(0);
		List<Row> rows = new ArrayList<Row>();
		sheet.forEach(row -> rows.add(row));

		// Remove header row
		rows.remove(0);

		// Initialize maps to store counts
		Map<Integer, Integer> yearCount = new HashMap<Integer, Integer>();
		Map<Integer, Map<Integer, Integer>> yearMonthCount = new HashMap<Integer, Map<Integer, Integer>>();
		Map<Integer, Map<Integer, Integer>> yearQuarterCount = new HashMap<Integer, Map<Integer, Integer>>();

		// Loop through rows to count approvals by year, month, and quarter
		for (Row row : rows) {

			// Get submission date and claim status
			Cell submissionDateCell = row.getCell(6);
			Cell claimStatusCell = row.getCell(7);

			if (submissionDateCell.getCellType() == CellType.NUMERIC
					&& DateUtil.isCellDateFormatted(submissionDateCell)) {

				// Get year and month from submission date

				Calendar cal = Calendar.getInstance();
				cal.setTime(submissionDateCell.getDateCellValue());

				int year = cal.get(Calendar.YEAR);
				int month = cal.get(Calendar.MONTH) + 1; // Add 1 to match with Excel month index

				// Update year count
				if (claimStatusCell.getStringCellValue().equalsIgnoreCase("Approved")) {
					yearCount.put(year, yearCount.getOrDefault(year, 0) + 1);

					// Update month count
					Map<Integer, Integer> monthCount = yearMonthCount.getOrDefault(year,
							new HashMap<Integer, Integer>());
					monthCount.put(month, monthCount.getOrDefault(month, 0) + 1);
					yearMonthCount.put(year, monthCount);

					// Update quarter count
					int quarter = (month - 1) / 3 + 1; // Calculate quarter based on month index
					Map<Integer, Integer> quarterCount = yearQuarterCount.getOrDefault(year,
							new HashMap<Integer, Integer>());
					quarterCount.put(quarter, quarterCount.getOrDefault(quarter, 0) + 1);
					yearQuarterCount.put(year, quarterCount);
				}
			}
		}

		Map<Integer, Integer> claimYearCount = new HashMap<Integer, Integer>();
		Map<Integer, Map<Integer, Integer>> claimYearMonthCount = new HashMap<Integer, Map<Integer, Integer>>();
		Map<Integer, Map<Integer, Integer>> claimYearQuarterCount = new HashMap<Integer, Map<Integer, Integer>>();

		// Loop through rows to count submissions by year, month, and quarter
		for (Row row : rows) {
			// Get submission date and claimed amount
			Cell submissionDateCell = row.getCell(6);
			Cell claimedAmountCell = row.getCell(5);
			if (submissionDateCell.getCellType() == CellType.NUMERIC
					&& DateUtil.isCellDateFormatted(submissionDateCell)) {

				// Get year and month from submission date
				Calendar cal = Calendar.getInstance();
				cal.setTime(submissionDateCell.getDateCellValue());

				int year = cal.get(Calendar.YEAR);
				int month = cal.get(Calendar.MONTH) + 1; // Add 1 to match with Excel month index

				// Update year count
				claimYearCount.put(year, claimYearCount.getOrDefault(year, 0) + 1);

				// Update month count
				Map<Integer, Integer> monthCount = claimYearMonthCount.getOrDefault(year,
						new HashMap<Integer, Integer>());
				monthCount.put(month, monthCount.getOrDefault(month, 0) + 1);
				claimYearMonthCount.put(year, monthCount);

				// Update quarter count
				int quarter = (month - 1) / 3 + 1; // Calculate quarter based on month index

				Map<Integer, Integer> quarterCount = claimYearQuarterCount.getOrDefault(year,
						new HashMap<Integer, Integer>());
				quarterCount.put(quarter, quarterCount.getOrDefault(quarter, 0) + 1);
				claimYearQuarterCount.put(year, quarterCount);
			}
		}

		// Print year count
		System.out.println("Count by year:");
		for (int year : yearCount.keySet()) {

			System.out.println(year + " Approved : " + yearCount.get(year));
			System.out.println(year + " Claimed: " + claimYearCount.get(year));
		}

		// Print month count
		System.out.println("Count by month:");
		for (int year : yearMonthCount.keySet()) {

			System.out.println("In " + year + ":");

			for (int month : yearMonthCount.get(year).keySet()) {
				System.out.println("month " + month + " Approved : " + yearMonthCount.get(year).get(month));
			}

			for (int month : claimYearMonthCount.get(year).keySet()) {
				System.out.println("Month " + month + " Claimed: " + claimYearMonthCount.get(year).get(month));
			}
		}

		// Print quarter count
		System.out.println("Count by quarter:");
		for (int year : claimYearCount.keySet()) {

			System.out.println(year + ": ");

			for (int quarter : yearQuarterCount.get(year).keySet()) {
				System.out.println("Quarter " + quarter + " Approved : " + yearQuarterCount.get(year).get(quarter));
			}

			for (int quarter : claimYearQuarterCount.get(year).keySet()) {
				System.out.println("Quarter " + quarter + " Claimed : " + claimYearQuarterCount.get(year).get(quarter));
			}
		}

		// Close workbook and input stream
		workbook.close();
		inputStream.close();
	}

	public static void getCategoryWiseClaims(String file) {
		Map<String, Double> categoryWiseTotalApprovedAmount = new HashMap<>();
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new File(file));
			Sheet sheet = workbook.getSheetAt(0);
			Row columnNamesRow = sheet.getRow(1);
			int claimCategoryIndex = -1;
			int approvedAmountIndex = -1;
			claimCategoryIndex = 1;
			approvedAmountIndex = 8;
			if (claimCategoryIndex == -1 || approvedAmountIndex == -1) {
				System.out.println("Either 'ClaimCategory' or 'ApprovedAmount' column is missing in the Excel sheet");
				return;
			}
			for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row.getCell(claimCategoryIndex) != null && row.getCell(approvedAmountIndex) != null) {
					String category = row.getCell(claimCategoryIndex).getStringCellValue();
					Double approvedAmount = row.getCell(approvedAmountIndex).getNumericCellValue();
					if (categoryWiseTotalApprovedAmount.containsKey(category)) {
						approvedAmount += categoryWiseTotalApprovedAmount.get(category);
					}
					categoryWiseTotalApprovedAmount.put(category, approvedAmount);
				}
			}
			for (String category : categoryWiseTotalApprovedAmount.keySet()) {
				System.out.println("Category: " + category + ", Total approved amount: "
						+ categoryWiseTotalApprovedAmount.get(category));
			}
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void claimProjections(String filePath) {
		int currentQuarter = 2; // Assuming the current quarter is Q2
		int totalQuarters = 4; // Assuming 4 quarters in a year
		int numPastQuarters = 3; // We want to calculate the average of the last 3 quarters
		Map<String, Double> categoryWisePastClaims = new HashMap<>();
		Map<String, Double> categoryWiseProjectedClaims = new HashMap<>();
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new File(filePath));
			Sheet sheet = workbook.getSheetAt(0);
			Row columnNamesRow = sheet.getRow(0);
			int categoryIndex = -1;
			int claimedAmountIndex = -1;
			categoryIndex = 1;
			claimedAmountIndex = 5;
			if (categoryIndex == -1 || claimedAmountIndex == -1) {
				System.out.println("Either 'Category' or 'ClaimedAmount' column is missing in the Excel sheet");
				return;
			}
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row.getCell(categoryIndex) != null && row.getCell(claimedAmountIndex) != null) {
					String category = row.getCell(categoryIndex).getStringCellValue();
					Double claimedAmount = row.getCell(claimedAmountIndex).getNumericCellValue();
					if (categoryWisePastClaims.containsKey(category)) {
						claimedAmount += categoryWisePastClaims.get(category);
					}
					categoryWisePastClaims.put(category, claimedAmount);
				}
			}
			for (String category : categoryWisePastClaims.keySet()) {
				Double pastQuarterClaims = categoryWisePastClaims.get(category);
				Double averagePastQuarterClaims = pastQuarterClaims / numPastQuarters;
				Double projectedQuarterClaims = averagePastQuarterClaims / currentQuarter * totalQuarters;
				categoryWiseProjectedClaims.put(category, projectedQuarterClaims);
			}
			for (String category : categoryWiseProjectedClaims.keySet()) {
				System.out.println("Category: " + category + ", Projected claims for next quarter: "
						+ new DecimalFormat("#.##").format(categoryWiseProjectedClaims.get(category)));
			}
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
