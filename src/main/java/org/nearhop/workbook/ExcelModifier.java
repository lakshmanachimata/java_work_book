/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package org.nearhop.workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelModifier {

    public static void main(String[] args) {
        String inputFilePath = "input.xlsx";   // Path to your existing Excel file
        String outputFilePath = "output.xlsx"; // Path where you want to save the modified file

        modifyExcelFile(inputFilePath, outputFilePath);
    }

    public static void modifyExcelFile(String inputFilePath, String outputFilePath) {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Access the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Modify cell in the first row, first column (A1)
            Row row = sheet.getRow(0);
            if (row == null) row = sheet.createRow(0); // Create row if it doesn't exist

            Cell cell = row.getCell(0);
            if (cell == null) cell = row.createCell(0); // Create cell if it doesn't exist

            cell.setCellValue("Hello, Excel!"); // Set new value in cell A1

            // Modify another cell, e.g., B2 (second row, second column)
            row = sheet.getRow(1); 
            if (row == null) row = sheet.createRow(1);

            cell = row.getCell(1);
            if (cell == null) cell = row.createCell(1);

            cell.setCellValue(12345678); // Set new integer value in cell B2

            // Save the changes to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
                System.out.println("Excel file modified and saved as " + outputFilePath);
            }

        } catch (IOException e) {
            System.err.println("Error modifying Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
}