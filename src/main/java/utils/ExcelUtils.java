package utils;

import org.apache.poi.ss.usermodel.*;  // Import for handling Excel sheets, rows, and cells
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  // Import for handling .xlsx Excel files

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

    // Represents the entire Excel workbook, encapsulates all sheets
    private static Workbook workbook;

    // Represents a specific sheet within the workbook
    private static Sheet sheet;

    // Path to the Excel file (to be reused when saving)
    private static String excelFilePath;

    /**
     * 1. Open the Excel file and access the specified sheet.
     * This is the first step before performing any read/write operations.
     *
     * @param filePath  Path to the Excel file
     * @param sheetName Name of the sheet to be accessed
     * @throws IOException If the file cannot be opened
     */
    public static void openExcelFile(String filePath, String sheetName) throws IOException {
        FileInputStream inputStream = new FileInputStream(filePath);  // Open the file input stream
        workbook = new XSSFWorkbook(inputStream);  // Load the workbook from the file
        sheet = workbook.getSheet(sheetName);  // Access the specific sheet by name
        excelFilePath = filePath;  // Store the file path for saving later
        inputStream.close();  // Close the input stream after loading the workbook
    }

    /**
     * 2. Write data to a specific cell.
     * This method allows you to modify any cell in the sheet (useful for updates).
     *
     * @param rowNum  Row number where data is to be written (0-based index)
     * @param colNum  Column number where data is to be written (0-based index)
     * @param value   The data to write into the specified cell
     */
    public static void setCellValue(int rowNum, int colNum, String value) {
        Row row = sheet.getRow(rowNum);  // Retrieve the specified row
        if (row == null) {
            row = sheet.createRow(rowNum);  // If the row doesn't exist, create it
        }
        Cell cell = row.getCell(colNum);  // Retrieve the specified cell
        if (cell == null) {
            cell = row.createCell(colNum);  // If the cell doesn't exist, create it
        }
        cell.setCellValue(value);  // Set the cell's value
    }

    /**
     * 3. Append data to the next empty row.
     * This method writes an entire array of data to the next available row (useful for adding new entries).
     *
     * @param values Array of values to write into the new row
     */
    public static void setCellValueInNextRow(String[] values) {
        int rowNum = sheet.getLastRowNum() + 1;  // Find the next empty row
        Row row = sheet.createRow(rowNum);  // Create a new row at that position

        // Loop through the values array and set each value in the corresponding column
        for (int colNum = 0; colNum < values.length; colNum++) {
            Cell cell = row.createCell(colNum);  // Create a new cell in the current column
            cell.setCellValue(values[colNum]);  // Set the value for the cell
        }
    }

    /**
     * 4. Save the workbook and close all file streams.
     * This method ensures that all changes are saved to the Excel file and frees up system resources.
     *
     * @throws IOException If there is an issue saving the file
     */
    public static void saveAndCloseExcel() throws IOException {
        FileOutputStream outputStream = new FileOutputStream(excelFilePath);  // Prepare to write changes to the file
        workbook.write(outputStream);  // Write the updated workbook to the file
        workbook.close();  // Close the workbook to free resources
        outputStream.close();  // Close the output stream to complete the save operation
    }

    /**
     * 5. Check if a row is empty.
     * This method helps to verify if a particular row is empty, useful before inserting data.
     *
     * @param rowNum Row number to check
     * @return true if the row is empty, false otherwise
     */
    public static boolean isRowEmpty(int rowNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return true;  // If row doesn't exist, it's considered empty
        }
        for (Cell cell : row) {
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;  // If any cell is not empty, the row isn't empty
            }
        }
        return true;  // If all cells are empty, the row is empty
    }
}