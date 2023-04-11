package chatgpt.test;

import java.io.IOException;
import java.util.Scanner;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		
		String filename = "C:\\Users\\hp\\Documents\\HackUseCase2Data.xlsx"; // replace with your file name
		ExcelValidator excelValidator = new ExcelValidator();		
		
		System.out.println("Starting to validate excel file at :"+filename);
		excelValidator.validateExcel(filename);
		System.out.println("Excel file validated\n");
		
		System.out.println("Start to evaluate the excel file !!!");
		ExcelEvaluator.evaluate(filename);
		
	}

}
