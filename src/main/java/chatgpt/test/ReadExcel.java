package chatgpt.test;

import java.io.IOException;

public class ReadExcel {

	public static void main(String[] args) throws IOException{
		
		String filename = "C:\\Users\\hp\\Documents\\HackUseCase2Data - Copy.xlsx"; // replace with your file name
		ExcelValidator excelValidator = new ExcelValidator();		
		
		System.out.println("Starting to validate excel file at :"+filename);
		try {
			excelValidator.validateExcel(filename);
		} catch (IOException e) {
			e.getMessage();
		} catch (ExcelException e) {
			System.out.println("Validation failed, message received :"+e.getMessage());
			return;
		}
		System.out.println("Excel file validated\n");
		
		System.out.println("Start to evaluate the excel file !!!");
		ExcelEvaluator.evaluate(filename);		
	}

}
