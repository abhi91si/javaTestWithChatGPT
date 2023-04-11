package chatgpt.test;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.DataFormatter;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class ExcelUtil {

	public static List<String> getColumnValueInList(FileInputStream inputStream, Sheet sheet, String columnname,
			boolean flagForDate) throws IOException {

		List<String> dates = new ArrayList<String>();
		List<String> values = new ArrayList<String>();

		int columnIndex = -1;
		Row headerRow = sheet.getRow(0);
		Iterator<Cell> headerCellIterator = headerRow.cellIterator();
		while (headerCellIterator.hasNext()) {
			Cell cell = headerCellIterator.next();
			String headerCellValue = cell.getStringCellValue();
			if (headerCellValue.equalsIgnoreCase(columnname)) {
				columnIndex = cell.getColumnIndex();
				break;
			}
		}

		Iterator<Row> rowIterator = sheet.rowIterator();
		rowIterator.next(); // Skip header row

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Cell cell = row.getCell(columnIndex);
			values = getvalueInList(cell, dates, flagForDate);
		}
		inputStream.close();

		return values;
	}

	public static List<String> getvalueInList(Cell cell, List<String> dates, boolean flag) {

		if (flag) {
			if (cell != null && cell.getCellType() != CellType.BLANK) {
				String cellValue;
				if (cell.getCellType() == CellType.NUMERIC) {
					SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
					cellValue = dateFormat.format(cell.getDateCellValue());
				} else {
					DataFormatter dataFormatter = new DataFormatter();
					cellValue = dataFormatter.formatCellValue(cell);
				}
				if (!cellValue.isEmpty()) {
					dates.add(cellValue);
				}
			}
		} else {
			if (cell != null) {
				String cellValue;
				if (cell.getCellType() == CellType.NUMERIC) {
					SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
					cellValue = dateFormat.format(cell.getDateCellValue());
				} else {
					DataFormatter dataFormatter = new DataFormatter();
					cellValue = dataFormatter.formatCellValue(cell);
				}
				dates.add(cellValue);
			}
		}
		return dates;
	}

}
