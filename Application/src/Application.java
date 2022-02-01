import java.io.IOException;

public class Application {

	public static void main(String[] args) throws IOException {
		WorkbookRead workbook = new WorkbookRead("Input.xlsx");
		WorkbookWrite updated_workbook = new WorkbookWrite();

		System.out.println("Started resolving");

		ParentChildResolver.resolve(workbook, updated_workbook, 7, 6);

		if (updated_workbook.writeToFile("result.xlsx")) {
			System.out.println("File Output Successfully");
		}

		workbook.close();
		System.out.print("End");
	}
}