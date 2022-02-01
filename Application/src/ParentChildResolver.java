import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ParentChildResolver {

	public static void resolve(WorkbookRead workbook, WorkbookWrite updated_workbook, int headerRow,
			int parentColumns) {
		Sheet readSheet = workbook.getSheet(0);
		Sheet writeSheet = updated_workbook.getWorkbook().createSheet();

		Iterator<Row> read_iterator = readSheet.iterator();
		int row_number;

		for (row_number = 0; row_number < headerRow; row_number++) {
			Row nextRow = read_iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			Row row = writeSheet.createRow((short) row_number);

			int column_number = 0;
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case STRING:
					String temp = cell.getStringCellValue();
					row.createCell(column_number++).setCellValue((String) temp);
					break;
				case NUMERIC:
					row.createCell(column_number++).setCellValue(cell.getNumericCellValue());
					break;
				}
			}
		}

		ArrayList<Object> currParent = new ArrayList<>();
		while (read_iterator.hasNext()) {
			Row nextRow = read_iterator.next();
			if (!isChild(nextRow)) {
				currParent = getParent(nextRow, parentColumns);
			} else {
				resolveParentChild(currParent, nextRow, writeSheet, row_number++);
			}
		}

	}

	public static boolean isChild(Row nextRow) {
		Cell cell = nextRow.getCell(0);
		if (cell.getCellType() == CellType.STRING && cell.getStringCellValue() == "") {
			return true;
		}

		return false;
	}

	public static void resolveParentChild(ArrayList<Object> parent_data, Row nextRow, Sheet spreadsheet,
			int row_number) {
		Iterator<Cell> cellIterator = nextRow.cellIterator();
		Row row = spreadsheet.createRow((short) row_number);

		int column_number = 0;

		for (Object field : parent_data) {
			Cell cell = row.createCell(column_number++);
			if (field instanceof String) {
				cell.setCellValue((String) field);
			} else if (field instanceof Double) {
				cell.setCellValue((Integer) field);
			}
		}

		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			switch (cell.getCellType()) {
			case STRING:
				String temp = cell.getStringCellValue();
				if (temp == "") {
					continue;
				}
				row.createCell(column_number++).setCellValue((String) temp);
				break;
			case NUMERIC:
				row.createCell(column_number++).setCellValue(cell.getNumericCellValue());
				break;
			}
		}
	}

	public static ArrayList<Object> getParent(Row nextRow, int ParentColumns) {
		Iterator<Cell> cellIterator = nextRow.cellIterator();
		ArrayList<Object> parentRowData = new ArrayList<>();

		int column_number = 0;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			switch (cell.getCellType()) {
			case STRING:
				String cellValue_s = cell.getStringCellValue();
				if (cellValue_s != "" || column_number < ParentColumns) {
					parentRowData.add(cell.getStringCellValue());
					column_number++;
				}
				break;
			case NUMERIC:
				double cellValue_d = cell.getNumericCellValue();
				parentRowData.add(cellValue_d);
				column_number++;
				break;
			}
		}
		return parentRowData;
	}

}
