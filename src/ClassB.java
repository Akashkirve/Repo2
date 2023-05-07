

	import java.io.*;
	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.*;

	public class ClassB {
	  public static void main(String[] args) throws Exception {
	    // Load the HTML file
	    File htmlFile = new File("C:\\Users\\HP\\NEW Workplace\\HTMLtoExcel\\src\\HTML File\\reports.html");
	    BufferedReader reader = new BufferedReader(new FileReader(htmlFile));
	    StringBuilder sb = new StringBuilder();
	    String line;
	    while ((line = reader.readLine()) != null) {
	      sb.append(line);
	    }
	    reader.close();
	    String htmlContent = sb.toString();

	    // Create an Excel workbook
	    XSSFWorkbook workbook = new XSSFWorkbook();
	    
	    
	    // Create a sheet
	    XSSFSheet sheet = workbook.createSheet("Report");

	    // Create a row for the header
	    XSSFRow headerRow = sheet.createRow(0);

	    // Create cells for the header
	    XSSFCell cell = headerRow.createCell(0);
	    cell.setCellValue("column 1");
	    cell = headerRow.createCell(1);
	    cell.setCellValue("Column 2");
	    cell = headerRow.createCell(2);
	    cell.setCellValue("Column 3");

	    // Parse the HTML and populate the sheet with data
	    // Use a HTML parsing library such as JSoup to extract data from the HTML
	    // and populate the sheet using the setCellValue() method

	    // Write the workbook to an Excel file
	    FileOutputStream outputStream = new FileOutputStream("C:\\Users\\HP\\NEW Workplacze\\HTMLtoExcel\\src\\HTML File\\report.xlsx");
	    workbook.write(outputStream);
	    workbook.close();
	    outputStream.close();
	  
	  }
	}




