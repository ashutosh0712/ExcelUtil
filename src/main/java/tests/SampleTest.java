package tests;  // Package declaration to group related classes

import utils.ExcelUtils;  // Import the ExcelUtils class for Excel operations
import java.io.IOException;  // Import IOException to handle input/output exceptions

public class SampleTest {
    public static void main(String[] args) {
        try {
            // Open the Excel file named "TestData.xlsx" and access the sheet "Sheet1"
            ExcelUtils.openExcelFile("TestData.xlsx", "Sheet1");

            //Arrays are defined to hold data (name and email). Each array represents a row of data to be inserted into the Excel sheet.
            // Define the data to be written into the first new row (user info)
            String[] data1 = { "Test User", "testuser@example.com","Hello" };

            // Define the data to be written into the second new row (user info)
            String[] data2 = { "New User", "newuser@example.com" };

            // Write the first set of data (data1) into the next available empty row
            ExcelUtils.setCellValueInNextRow(data1);

            // Write the second set of data (data2) into the next available empty row
            ExcelUtils.setCellValueInNextRow(data2);

            // Save the changes and close the Excel file
            ExcelUtils.saveAndCloseExcel();

            // Print a confirmation message to the console
            System.out.println("Data entered successfully");
        } catch (IOException e) {
            // Catch any IOExceptions that might occur during file operations
            e.printStackTrace();  // Print the stack trace to help with debugging
        }
    }
}