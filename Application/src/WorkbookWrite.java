import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookWrite {

	private XSSFWorkbook workbook;

	public WorkbookWrite() {

		workbook = new XSSFWorkbook();
	}

	public XSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	public boolean writeToFile(String fileName) {
		try {
			FileOutputStream out = new FileOutputStream(new File(fileName));
			this.workbook.write(out);
			out.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
}
