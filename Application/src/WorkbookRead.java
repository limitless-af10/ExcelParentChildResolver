
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;


public class WorkbookRead {
	
	private FileInputStream inputStream;
	private Workbook workbook; 
	
	public WorkbookRead(String filePath) throws IOException{
		inputStream = new FileInputStream(new File(filePath));
		workbook = new XSSFWorkbook(inputStream);
		
	}
	
	public Workbook getWorkbook() {
		return this.workbook;
	}
	
	public void close() throws IOException{
		workbook.close();
        inputStream.close();
	}
	
	public Sheet getSheet(int sheet_number) {
		return this.workbook.getSheetAt(sheet_number);
	}

}
