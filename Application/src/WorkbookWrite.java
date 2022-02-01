import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WorkbookWrite {

	private HSSFWorkbook workbook;

	public WorkbookWrite() {

		workbook = new HSSFWorkbook();
	}

	public HSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	public boolean writeToFile(String fileName) {
		try {
			FileOutputStream out = new FileOutputStream(new File(fileName));
			this.workbook.write(out);
			out.close();
			workbook.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
}
