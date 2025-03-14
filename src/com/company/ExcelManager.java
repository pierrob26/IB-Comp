package com.company;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.List;

public class ExcelManager {

    public static void createExcelFile(String filename) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream(filename);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
        System.out.println("Excel file '" + filename + "' created successfully.");
    }

    public static void addSheet(String filename, String sheetName) throws IOException {
        FileInputStream file = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        workbook.createSheet(sheetName);
        file.close();
        FileOutputStream outFile = new FileOutputStream(filename);
        workbook.write(outFile);
        outFile.close();
        workbook.close();
        System.out.println("Sheet '" + sheetName + "' added to '" + filename + "'.");
    }

    public static void writeToCell(String filename, String sheetName, int row, int col, String value) throws IOException {
        FileInputStream file = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        Row rowObj = sheet.getRow(row);
        if (rowObj == null) {
            rowObj = sheet.createRow(row);
        }
        Cell cell = rowObj.createCell(col);
        cell.setCellValue(value);
        file.close();
        FileOutputStream outFile = new FileOutputStream(filename);
        workbook.write(outFile);
        outFile.close();
        workbook.close();
        System.out.println("Written '" + value + "' to cell(" + row + ", " + col + ") in sheet '" + sheetName + "'.");
    }

    public static String readFromCell(String filename, String sheetName, int row, int col) throws IOException {
        FileInputStream file = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(sheetName);
        String value = "";
        if (sheet != null) {
            Row rowObj = sheet.getRow(row);
            if (rowObj != null) {
                Cell cell = rowObj.getCell(col);
                if (cell != null) {
                    value = cell.getStringCellValue();
                }
            }
        }
        file.close();
        workbook.close();
        System.out.println("Value at cell(" + row + ", " + col + ") in sheet '" + sheetName + "': " + value);
        return value;
    }

    public static void appendRow(String filename, String sheetName, List<String> rowData) throws IOException {
        FileInputStream file = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        int lastRow = sheet.getLastRowNum() + 1;
        Row row = sheet.createRow(lastRow);
        for (int i = 0; i < rowData.size(); i++) {
            row.createCell(i).setCellValue(rowData.get(i));
        }
        file.close();
        FileOutputStream outFile = new FileOutputStream(filename);
        workbook.write(outFile);
        outFile.close();
        workbook.close();
        System.out.println("Appended row " + rowData + " to sheet '" + sheetName + "'.");
    }
}
